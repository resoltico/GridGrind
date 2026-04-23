package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

/** Builds contract-owned starter request scaffolds from task descriptors. */
public final class GridGrindTaskPlanner {
  private GridGrindTaskPlanner() {}

  /** Returns a starter task-plan scaffold for one stable task id. */
  public static TaskPlanTemplate templateFor(String taskId) {
    return planFor(
        GridGrindTaskCatalog.entryFor(taskId)
            .orElseThrow(
                () ->
                    new IllegalArgumentException(
                        "Unknown task id for task planning: "
                            + CatalogRecordValidation.requireNonBlank(taskId, "taskId"))));
  }

  /** Returns a starter task-plan scaffold for one task entry. */
  public static TaskPlanTemplate planFor(TaskEntry task) {
    TaskEntry taskEntry = java.util.Objects.requireNonNull(task, "task must not be null");
    WorkbookPlan.WorkbookSource source = sourceFor(taskEntry);
    WorkbookPlan.WorkbookPersistence persistence = persistenceFor(taskEntry, source);
    WorkbookPlan requestTemplate = new WorkbookPlan(source, persistence, List.of());
    return new TaskPlanTemplate(
        GridGrindProtocolVersion.current(), taskEntry, requestTemplate, authoringNotes(taskEntry));
  }

  private static WorkbookPlan.WorkbookSource sourceFor(TaskEntry task) {
    return switch (soleCapabilityId(task, "sourceTypes")) {
      case "NEW" -> new WorkbookPlan.WorkbookSource.New();
      case "EXISTING" -> new WorkbookPlan.WorkbookSource.ExistingFile(defaultInputPath(task));
      default ->
          throw new IllegalStateException(
              "Task " + task.id() + " references unsupported source type for planning");
    };
  }

  private static WorkbookPlan.WorkbookPersistence persistenceFor(
      TaskEntry task, WorkbookPlan.WorkbookSource source) {
    String persistenceType = soleCapabilityId(task, "persistenceTypes");
    return switch (persistenceType) {
      case "NONE" -> new WorkbookPlan.WorkbookPersistence.None();
      case "SAVE_AS" -> new WorkbookPlan.WorkbookPersistence.SaveAs(defaultOutputPath(task));
      case "OVERWRITE" -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile)) {
          throw new IllegalStateException(
              "Task "
                  + task.id()
                  + " cannot plan OVERWRITE persistence without an EXISTING source");
        }
        yield new WorkbookPlan.WorkbookPersistence.OverwriteSource();
      }
      default ->
          throw new IllegalStateException(
              "Task " + task.id() + " references unsupported persistence type for planning");
    };
  }

  private static String soleCapabilityId(TaskEntry task, String group) {
    Set<String> ids = capabilityIds(task, group);
    if (ids.isEmpty()) {
      throw new IllegalStateException(
          "Task " + task.id() + " does not declare any " + group + " capability");
    }
    if (ids.size() > 1) {
      throw new IllegalStateException(
          "Task " + task.id() + " declares multiple " + group + " capabilities: " + ids);
    }
    return ids.iterator().next();
  }

  private static Set<String> capabilityIds(TaskEntry task, String group) {
    Set<String> ids = new LinkedHashSet<>();
    for (TaskPhase phase : task.phases()) {
      for (TaskCapabilityRef capabilityRef : phase.capabilityRefs()) {
        if (capabilityRef.group().equals(group)) {
          ids.add(capabilityRef.id());
        }
      }
    }
    return ids;
  }

  private static List<String> authoringNotes(TaskEntry task) {
    List<String> baseNotes =
        List.of(
            "requestTemplate is intentionally minimal and valid: source and persistence are"
                + " scaffolded, but steps stays empty until you author the workflow.",
            "Use task.phases[*].capabilityRefs to discover the exact operation shapes through"
                + " --print-protocol-catalog --search <text> or"
                + " --print-protocol-catalog --operation <group>:<id>.",
            "Replace any TODO-style .xlsx placeholder path before execution.");
    if (capabilityIds(task, "persistenceTypes").contains("NONE")) {
      return List.of(
          baseNotes.get(0),
          baseNotes.get(1),
          "This task defaults to in-memory persistence so discovery and inspection stay"
              + " non-destructive.");
    }
    return baseNotes;
  }

  private static String defaultInputPath(TaskEntry task) {
    return "todo-" + slug(task.id()) + "-input.xlsx";
  }

  private static String defaultOutputPath(TaskEntry task) {
    return "todo-" + slug(task.id()) + "-output.xlsx";
  }

  private static String slug(String taskId) {
    return taskId.toLowerCase(Locale.ROOT).replace('_', '-');
  }
}
