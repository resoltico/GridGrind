package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskCapabilityRef;
import dev.erst.gridgrind.cli.discovery.TaskDiscoveryProfile;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskExecutionProfile;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskIntentProfile;
import dev.erst.gridgrind.cli.discovery.TaskInteractionProfile;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskNarrative;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhase;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskStarterContract;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import dev.erst.gridgrind.cli.discovery.TaskWorkflow;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Shared builders for the canonical published task-recipe definitions. */
final class GridGrindTaskRecipeSupport {
  private GridGrindTaskRecipeSupport() {}

  static GridGrindTaskRecipeDefinition definition(TaskEntry task, WorkbookPlan starterPlan) {
    return new GridGrindTaskRecipeDefinition(task, starterPlan);
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
    return TaskStarterContract.selfContained(TaskStarterRecipeSupport.taskRequestFileName(taskId));
  }

  static TaskStarterContract assetBackedStarter(String taskId, String... requiredWorkspacePaths) {
    return TaskStarterContract.assetBacked(
        TaskStarterRecipeSupport.taskRequestFileName(taskId), requiredWorkspacePaths);
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
