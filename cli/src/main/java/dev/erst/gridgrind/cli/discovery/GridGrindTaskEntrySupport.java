package dev.erst.gridgrind.cli.discovery;

import java.util.List;

/** Shared builders for the CLI-owned task descriptor registry. */
final class GridGrindTaskEntrySupport {
  private GridGrindTaskEntrySupport() {}

  static TaskExecutionProfile profile(
      TaskSourceMode sourceMode,
      TaskPersistenceMode persistenceMode,
      TaskMutationMode mutationMode,
      TaskAssetMode assetMode) {
    return new TaskExecutionProfile(sourceMode, persistenceMode, mutationMode, assetMode);
  }

  static TaskEntry task(
      String id,
      TaskDiscoveryProfile discoveryProfile,
      TaskNarrative narrative,
      TaskExecutionProfile executionProfile,
      TaskInteractionProfile interactionProfile,
      TaskWorkflow workflow) {
    return new TaskEntry(
        id, discoveryProfile, narrative, executionProfile, interactionProfile, workflow);
  }

  static TaskDiscoveryProfile discovery(
      List<String> discoveryTerms, List<String> intentTags, TaskIntentProfile intentProfile) {
    return new TaskDiscoveryProfile(discoveryTerms, intentTags, intentProfile);
  }

  static TaskNarrative narrative(
      String summary,
      List<String> outcomes,
      List<String> requiredInputs,
      List<String> optionalFeatures) {
    return new TaskNarrative(summary, outcomes, requiredInputs, optionalFeatures);
  }

  static TaskIntentProfile intent(List<TaskGoalKind> goals, List<TaskArtifactKind> artifacts) {
    return new TaskIntentProfile(goals, artifacts);
  }

  static TaskInteractionProfile signals(
      List<TaskInputKind> requiredInputKinds, List<TaskVerificationKind> verificationKinds) {
    return new TaskInteractionProfile(requiredInputKinds, verificationKinds);
  }

  static TaskWorkflow workflow(List<TaskPhase> phases, List<String> commonPitfalls) {
    return new TaskWorkflow(phases, commonPitfalls);
  }

  static TaskPhase phase(
      TaskPhasePurpose purpose,
      String label,
      String objective,
      List<TaskCapabilityRef> capabilityRefs,
      List<String> notes) {
    return new TaskPhase(purpose, label, objective, capabilityRefs, notes);
  }

  static TaskCapabilityRef ref(String group, String id) {
    return new TaskCapabilityRef(group, id);
  }
}
