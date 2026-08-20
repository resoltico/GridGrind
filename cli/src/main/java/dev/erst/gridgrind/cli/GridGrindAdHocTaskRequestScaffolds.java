package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.TaskCapabilityRef;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskExecutionProfile;
import dev.erst.gridgrind.cli.discovery.TaskPhase;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.ProtocolStepTemplate;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;

/** Builds generic ad hoc request scaffolds for unpublished task descriptors only. */
final class GridGrindAdHocTaskRequestScaffolds {
  private GridGrindAdHocTaskRequestScaffolds() {}

  /** Returns one generic executable starter request for one unpublished task descriptor. */
  static WorkbookPlan requestFor(TaskEntry task) {
    TaskEntry taskEntry = java.util.Objects.requireNonNull(task, "task must not be null");
    if (GridGrindCliRecipeRegistry.recipeFor(taskEntry.id()).isPresent()) {
      throw new IllegalArgumentException(
          "Published recipe id " + taskEntry.id() + " must use the canonical recipe registry");
    }
    TaskExecutionProfile profile = taskEntry.executionProfile();
    WorkbookPlan.WorkbookSource source = sourceFor(taskEntry.id(), profile);
    WorkbookPlan.WorkbookPersistence persistence = persistenceFor(taskEntry, source);
    WorkbookPlan templatePlan =
        WorkbookPlan.standard(
            source,
            persistence,
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    return decodedRequest(taskEntry.id(), requestTemplateFor(taskEntry, templatePlan));
  }

  private static ObjectNode requestTemplateFor(TaskEntry task, WorkbookPlan templatePlan) {
    ObjectNode requestTemplate =
        dev.erst.gridgrind.contract.json.GridGrindJsonOutput.requestTree(templatePlan);
    requestTemplate.set("steps", stepsFor(task));
    return requestTemplate;
  }

  private static ArrayNode stepsFor(TaskEntry task) {
    ArrayNode steps = tools.jackson.databind.node.JsonNodeFactory.instance.arrayNode();
    Set<String> emittedStepBodies = new LinkedHashSet<>();
    int phaseIndex = 1;
    for (TaskPhase phase : task.workflow().phases()) {
      int emittedStepIndex = 1;
      for (TaskCapabilityRef capabilityRef : phase.capabilityRefs()) {
        int emittedPhaseIndex = phaseIndex;
        java.util.Optional<ProtocolStepTemplate> optionalStepTemplate =
            GridGrindProtocolCatalog.entryFor(capabilityRef.qualifiedId())
                .flatMap(entry -> entry.stepTemplate());
        if (optionalStepTemplate.isPresent()) {
          ProtocolStepTemplate stepTemplate = optionalStepTemplate.orElseThrow();
          String stepBodyKey = stepTemplate.stepKind() + ":" + stepTemplate.bodyField();
          if (!emittedStepBodies.add(stepBodyKey + ":" + capabilityRef.qualifiedId())) {
            continue;
          }
          steps.add(renderedStep(stepTemplate, emittedPhaseIndex, emittedStepIndex));
          emittedStepIndex++;
        }
      }
      phaseIndex++;
    }
    return steps;
  }

  private static ObjectNode renderedStep(
      ProtocolStepTemplate template, int phaseIndex, int capabilityIndex) {
    ObjectNode step = template.template().deepCopy();
    String stepId = "phase-" + phaseIndex + "-step-" + capabilityIndex;
    step.put("stepId", stepId);
    return step;
  }

  private static WorkbookPlan.WorkbookSource sourceFor(
      String taskId, TaskExecutionProfile executionProfile) {
    return switch (executionProfile.sourceMode()) {
      case NEW_WORKBOOK -> new WorkbookPlan.WorkbookSource.New();
      case EXISTING_WORKBOOK ->
          new WorkbookPlan.WorkbookSource.ExistingFile(defaultInputPath(taskId));
    };
  }

  private static WorkbookPlan.WorkbookPersistence persistenceFor(
      TaskEntry task, WorkbookPlan.WorkbookSource source) {
    return switch (task.executionProfile().persistenceMode()) {
      case NONE -> new WorkbookPlan.WorkbookPersistence.None();
      case SAVE_AS -> {
        String outputPath = defaultOutputPath(task.id());
        yield source instanceof WorkbookPlan.WorkbookSource.ExistingFile
            ? new WorkbookPlan.WorkbookPersistence.SaveAs(
                outputPath,
                WorkbookPlan.WorkbookPersistence.IfExists.REPLACE,
                dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none())
            : new WorkbookPlan.WorkbookPersistence.SaveAs(
                outputPath, WorkbookPlan.WorkbookPersistence.IfExists.REPLACE);
      }
      case OVERWRITE -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile)) {
          throw new IllegalStateException(
              "Task "
                  + task.id()
                  + " cannot plan OVERWRITE persistence without an EXISTING source");
        }
        yield new WorkbookPlan.WorkbookPersistence.Overwrite(
            dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none());
      }
    };
  }

  static WorkbookPlan decodedRequest(String taskId, ObjectNode requestTemplate) {
    try {
      ByteArrayOutputStream buffer = new ByteArrayOutputStream();
      GridGrindJsonOutput.writeCatalogLookupValue(buffer, requestTemplate);
      return GridGrindJson.readRequest(buffer.toByteArray());
    } catch (IOException | RuntimeException exception) {
      throw new IllegalStateException(
          "CLI-generated task request is invalid for " + requireNonBlank(taskId, "taskId"),
          exception);
    }
  }

  private static String defaultInputPath(String taskId) {
    return "starter-" + slug(taskId) + "-input.xlsx";
  }

  private static String defaultOutputPath(String taskId) {
    return "starter-" + slug(taskId) + "-output.xlsx";
  }

  private static String slug(String taskId) {
    return taskId.toLowerCase(Locale.ROOT).replace('_', '-');
  }

  private static String requireNonBlank(String value, String fieldName) {
    java.util.Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
