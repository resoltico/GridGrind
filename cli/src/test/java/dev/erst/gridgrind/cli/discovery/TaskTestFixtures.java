package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Locale;

/** Shared builders for tests that need compact ad hoc task descriptors. */
public final class TaskTestFixtures {
  private TaskTestFixtures() {}

  public static TaskEntry task(
      String id, TaskExecutionProfile executionProfile, List<TaskPhase> phases) {
    return new TaskEntry(
        id,
        List.of("office"),
        discoveryProfile(id),
        narrative("summary"),
        executionProfile,
        interactionProfile(),
        TaskStarterContract.selfContained(
            id.toLowerCase(Locale.ROOT).replace('_', '-') + "-request.json"),
        workflow(phases));
  }

  public static TaskDiscoveryProfile discoveryProfile(String id) {
    String token = id.toLowerCase(Locale.ROOT).replace('_', ' ');
    return new TaskDiscoveryProfile(
        List.of(token),
        new TaskIntentProfile(List.of(TaskGoalKind.AUTHOR), List.of(TaskArtifactKind.WORKBOOK)));
  }

  public static TaskNarrative narrative(String summary) {
    return new TaskNarrative(summary, List.of("outcome"), List.of("input"), List.of("feature"));
  }

  public static TaskInteractionProfile interactionProfile() {
    return new TaskInteractionProfile(List.of(), List.of());
  }

  public static TaskWorkflow workflow(List<TaskPhase> phases) {
    return new TaskWorkflow(phases, List.of("pitfall"));
  }

  public static TaskPhase phase(List<TaskCapabilityRef> capabilityRefs) {
    return new TaskPhase(
        TaskPhasePurpose.AUTHOR, "Phase", "Objective", capabilityRefs, List.of("note"));
  }

  public static TaskExecutionProfile profile() {
    return new TaskExecutionProfile(
        TaskSourceMode.NEW_WORKBOOK,
        TaskPersistenceMode.NONE,
        TaskMutationMode.MUTATING,
        TaskAssetMode.SELF_CONTAINED);
  }
}
