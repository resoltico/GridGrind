package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskExecutionProfile;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPlanTemplate;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/** Builds CLI-owned starter request scaffolds from public task descriptors. */
final class GridGrindTaskPlanner {
  private GridGrindTaskPlanner() {}

  /** Returns a starter task-plan scaffold for one stable task id. */
  static TaskPlanTemplate templateFor(String taskId) {
    String requestedTaskId = requireNonBlank(taskId, "taskId");
    return GridGrindTaskCatalog.entryFor(requestedTaskId)
        .map(GridGrindTaskPlanner::planFor)
        .orElseThrow(
            () ->
                new IllegalArgumentException(
                    "Unknown task id for task planning: " + requestedTaskId));
  }

  /** Returns a starter task-plan scaffold for one task entry. */
  static TaskPlanTemplate planFor(TaskEntry task) {
    TaskEntry taskEntry = java.util.Objects.requireNonNull(task, "task must not be null");
    TaskExecutionProfile profile = taskEntry.executionProfile();
    WorkbookPlan.WorkbookSource source = sourceFor(taskEntry.id(), profile);
    WorkbookPlan.WorkbookPersistence persistence = persistenceFor(taskEntry, source);
    WorkbookPlan requestTemplate =
        WorkbookPlan.standard(
            source,
            persistence,
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    return new TaskPlanTemplate(
        GridGrindProtocolVersion.current(), taskEntry, requestTemplate, authoringNotes(taskEntry));
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
      case SAVE_AS -> new WorkbookPlan.WorkbookPersistence.SaveAs(defaultOutputPath(task.id()));
      case OVERWRITE_SOURCE -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile)) {
          throw new IllegalStateException(
              "Task "
                  + task.id()
                  + " cannot plan OVERWRITE persistence without an EXISTING source");
        }
        yield new WorkbookPlan.WorkbookPersistence.OverwriteSource();
      }
    };
  }

  private static List<String> authoringNotes(TaskEntry task) {
    List<String> notes = new ArrayList<>();
    notes.add(
        "requestTemplate is intentionally minimal for this task: source and persistence are"
            + " scaffolded, and you author the exact workflow steps from the task phases.");
    notes.add(
        "Use task.phases[*].capabilityRefs to discover the exact operation shapes through"
            + " --print-protocol-catalog --search <text> or"
            + " --print-protocol-catalog --operation <group>:<id>.");
    if (task.executionProfile().sourceMode() == TaskSourceMode.EXISTING_WORKBOOK) {
      notes.add("Replace the placeholder input workbook path before execution.");
    }
    if (task.executionProfile().persistenceMode() == TaskPersistenceMode.SAVE_AS) {
      notes.add("Replace the placeholder output workbook path before execution.");
    }
    if (task.executionProfile().persistenceMode() == TaskPersistenceMode.NONE) {
      notes.add(
          "This task defaults to in-memory persistence so discovery and inspection stay"
              + " non-destructive.");
    }
    if (task.executionProfile().assetMode() == TaskAssetMode.REQUIRES_EXTERNAL_PAYLOADS) {
      notes.add(
          "This task expects external payload files or mapped workbook assets; replace the"
              + " placeholder sources before execution.");
    }
    for (String pitfall : task.commonPitfalls()) {
      notes.add("Common pitfall: " + pitfall);
    }
    return List.copyOf(notes);
  }

  private static String defaultInputPath(String taskId) {
    return "todo-" + slug(taskId) + "-input.xlsx";
  }

  private static String defaultOutputPath(String taskId) {
    return "todo-" + slug(taskId) + "-output.xlsx";
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
