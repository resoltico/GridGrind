package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** Detailed lookup payload for one published task-starter recipe. */
@Variant(
    id = "TASK_STARTER",
    summary = "Detailed lookup payload for one published task-starter recipe.")
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
    id = TextValidation.requireNonBlank(id, "id");
    requestFileName = RecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = TextValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        RecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    RecipeCatalogValidation.validateWorkspaceContract(advisory, requiredWorkspacePaths, "recipes");
    intentTags = DiscoveryValidation.copyStrings(intentTags, "intentTags");
    Objects.requireNonNull(discoveryProfile, "discoveryProfile must not be null");
    outcomes = DiscoveryValidation.copyStrings(outcomes, "outcomes");
    requiredInputs = DiscoveryValidation.copyStrings(requiredInputs, "requiredInputs");
    optionalFeatures = DiscoveryValidation.copyStrings(optionalFeatures, "optionalFeatures");
    Objects.requireNonNull(executionProfile, "executionProfile must not be null");
    Objects.requireNonNull(interactionProfile, "interactionProfile must not be null");
    Objects.requireNonNull(workflow, "workflow must not be null");
    Objects.requireNonNull(requestProfile, "requestProfile must not be null");
  }
}
