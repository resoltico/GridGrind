package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** Detailed lookup payload for one CLI-authored task-starter recipe. */
public record TaskStarterRecipeCatalogDetail(
    String id,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    List<String> requiredWorkspacePaths,
    List<String> intentTags,
    TaskDiscoveryProfile discoveryProfile,
    List<String> outcomes,
    List<String> requiredInputs,
    List<String> optionalFeatures,
    TaskExecutionProfile executionProfile,
    TaskInteractionProfile interactionProfile,
    TaskWorkflow workflow,
    RecipeRequestProfile requestProfile)
    implements RecipeCatalogDetail {
  public TaskStarterRecipeCatalogDetail {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    requestFileName = CliRecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        CliRecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    CliRecipeCatalogValidation.validateWorkspaceContract(
        advisory, requiredWorkspacePaths, "recipes");
    intentTags = CliDiscoveryValidation.copyStrings(intentTags, "intentTags");
    Objects.requireNonNull(discoveryProfile, "discoveryProfile must not be null");
    outcomes = CliDiscoveryValidation.copyStrings(outcomes, "outcomes");
    requiredInputs = CliDiscoveryValidation.copyStrings(requiredInputs, "requiredInputs");
    optionalFeatures = CliDiscoveryValidation.copyStrings(optionalFeatures, "optionalFeatures");
    Objects.requireNonNull(executionProfile, "executionProfile must not be null");
    Objects.requireNonNull(interactionProfile, "interactionProfile must not be null");
    Objects.requireNonNull(workflow, "workflow must not be null");
    Objects.requireNonNull(requestProfile, "requestProfile must not be null");
  }
}
