package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.List;

/** Shared builders for the canonical published task-recipe definitions. */
final class TaskRecipes {
  private TaskRecipes() {}

  static TaskRecipeTemplate definition(TaskEntry task, WorkbookPlan starterPlan) {
    return new TaskRecipeTemplate(task, starterPlan);
  }

  static TaskExecutionProfile profile(
      TaskSourceMode sourceMode,
      TaskPersistenceMode persistenceMode,
      TaskMutationMode mutationMode,
      TaskAssetMode assetMode) {
    return new TaskExecutionProfile(sourceMode, persistenceMode, mutationMode, assetMode);
  }

  static TaskEntry task(
      String id,
      TaskStarterContract starter,
      List<String> intentTags,
      TaskDiscoveryProfile discoveryProfile,
      TaskNarrative narrative,
      TaskExecutionProfile executionProfile,
      TaskInteractionProfile interactionProfile,
      TaskWorkflow workflow) {
    return new TaskEntry(
        id,
        intentTags,
        discoveryProfile,
        narrative,
        executionProfile,
        interactionProfile,
        starter,
        workflow);
  }

  static TaskStarterContract selfContainedStarter(String taskId) {
    return TaskStarterContract.selfContained(TaskStarters.taskRequestFileName(taskId));
  }

  static TaskStarterContract assetBackedStarter(String taskId, String... requiredWorkspacePaths) {
    return TaskStarterContract.assetBacked(
        TaskStarters.taskRequestFileName(taskId), requiredWorkspacePaths);
  }

  static TaskDiscoveryProfile discovery(
      List<String> discoveryTerms, TaskIntentProfile intentProfile) {
    return new TaskDiscoveryProfile(discoveryTerms, intentProfile);
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
