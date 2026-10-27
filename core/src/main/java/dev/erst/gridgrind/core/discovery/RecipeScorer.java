package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Scores built-in recipes against one normalized English keyword query. */
final class RecipeScorer {
  private RecipeScorer() {}

  static Optional<RecipeKeywordMatchReport.Candidate> candidateFor(
      Recipe recipe, List<String> queryTerms) {
    Objects.requireNonNull(recipe, "recipe must not be null");
    return switch (recipe.type()) {
      case EXAMPLE -> exampleCandidateFor(recipe, queryTerms);
      case TASK_STARTER -> taskCandidateFor(recipe, queryTerms);
    };
  }

  private static Optional<RecipeKeywordMatchReport.Candidate> exampleCandidateFor(
      Recipe recipe, List<String> queryTerms) {
    MatchAccumulator accumulator = new MatchAccumulator();
    scoreSurface(
        queryTerms, List.of(recipe.id().replace('_', ' ')), "recipe id", 14, true, accumulator);
    scoreSurface(queryTerms, recipe.intentTags(), "intent tag", 12, true, accumulator);
    scoreSurface(
        queryTerms,
        List.of(exampleStem(recipe.requestFileName())),
        "file stem",
        8,
        true,
        accumulator);
    scoreSurface(queryTerms, List.of(recipe.summary()), "summary", 7, true, accumulator);
    scoreSurface(
        queryTerms, recipe.requiredWorkspacePaths(), "required asset", 2, false, accumulator);
    if (accumulator.score == 0 || !accumulator.hasSemanticMatch) {
      return Optional.empty();
    }
    return Optional.of(
        new RecipeKeywordMatchReport.Candidate(
            recipe.id(),
            RecipeType.EXAMPLE,
            recipe.summary(),
            accumulator.score,
            List.copyOf(accumulator.matchedTerms),
            accumulator.compactMatchSources()));
  }

  private static Optional<RecipeKeywordMatchReport.Candidate> taskCandidateFor(
      Recipe recipe, List<String> queryTerms) {
    TaskEntry task = RecipeRegistry.taskEntryFor(recipe.id()).orElseThrow();
    MatchAccumulator accumulator = new MatchAccumulator();
    scoreSurface(
        queryTerms, List.of(task.id().replace('_', ' ')), "recipe id", 14, true, accumulator);
    scoreSurface(queryTerms, task.intentTags(), "intent tag", 11, true, accumulator);
    scoreDiscoveryProfile(queryTerms, task.discoveryProfile(), accumulator);
    scoreTypedTaskSurface(
        queryTerms,
        task.discoveryProfile().intentProfile().goals(),
        task.discoveryProfile().intentProfile().artifacts(),
        task.interactionProfile().requiredInputKinds(),
        task.interactionProfile().verificationKinds(),
        accumulator);
    scoreSurface(queryTerms, List.of(task.narrative().summary()), "summary", 6, true, accumulator);
    scoreSurface(queryTerms, task.narrative().outcomes(), "outcome", 4, true, accumulator);
    scoreSurface(
        queryTerms, task.narrative().requiredInputs(), "required input", 3, false, accumulator);
    scoreSurface(
        queryTerms, task.narrative().optionalFeatures(), "optional feature", 2, false, accumulator);
    scoreExecutionProfileSurface(queryTerms, task.executionProfile(), accumulator);
    scorePhaseSurface(queryTerms, task.workflow().phases(), accumulator);
    scoreCapabilitySurface(queryTerms, task.workflow().phases(), accumulator);
    if (accumulator.score == 0 || !accumulator.hasSemanticMatch) {
      return Optional.empty();
    }
    return Optional.of(
        new RecipeKeywordMatchReport.Candidate(
            task.id(),
            RecipeType.TASK_STARTER,
            task.narrative().summary(),
            accumulator.score,
            List.copyOf(accumulator.matchedTerms),
            accumulator.compactMatchSources()));
  }

  private static void scoreDiscoveryProfile(
      List<String> queryTerms,
      TaskDiscoveryProfile discoveryProfile,
      MatchAccumulator accumulator) {
    scoreSurface(
        queryTerms, discoveryProfile.discoveryTerms(), "discovery term", 13, true, accumulator);
  }

  private static void scoreExecutionProfileSurface(
      List<String> queryTerms, TaskExecutionProfile profile, MatchAccumulator accumulator) {
    scoreSurface(
        queryTerms,
        List.of(TaskKeywords.sourceModeSurface(profile.sourceMode())),
        "source mode",
        6,
        false,
        accumulator);
    scoreSurface(
        queryTerms,
        List.of(TaskKeywords.persistenceModeSurface(profile.persistenceMode())),
        "persistence mode",
        6,
        false,
        accumulator);
    scoreSurface(
        queryTerms,
        List.of(TaskKeywords.mutationModeSurface(profile.mutationMode())),
        "mutation mode",
        6,
        false,
        accumulator);
    scoreSurface(
        queryTerms,
        List.of(TaskKeywords.assetModeSurface(profile.assetMode())),
        "asset mode",
        5,
        false,
        accumulator);
  }

  private static void scorePhaseSurface(
      List<String> queryTerms, List<TaskPhase> phases, MatchAccumulator accumulator) {
    for (TaskPhase phase : phases) {
      scoreSurface(
          queryTerms,
          List.of(TaskKeywords.phasePurposeSurface(phase.purpose())),
          "phase purpose",
          4,
          true,
          accumulator);
      scoreSurface(queryTerms, List.of(phase.label()), "phase label", 2, true, accumulator);
      scoreSurface(
          queryTerms, List.of(phase.objective()), "phase objective", 1, false, accumulator);
    }
  }

  private static void scoreTypedTaskSurface(
      List<String> queryTerms,
      List<TaskGoalKind> goals,
      List<TaskArtifactKind> artifacts,
      List<TaskInputKind> requiredInputKinds,
      List<TaskVerificationKind> verificationKinds,
      MatchAccumulator accumulator) {
    for (TaskGoalKind goal : goals) {
      scoreSurface(
          queryTerms, List.of(TaskKeywords.goalSurface(goal)), "task goal", 10, true, accumulator);
    }
    for (TaskArtifactKind artifact : artifacts) {
      scoreSurface(
          queryTerms,
          List.of(TaskKeywords.artifactSurface(artifact)),
          "artifact",
          3,
          false,
          accumulator);
    }
    for (TaskInputKind inputKind : requiredInputKinds) {
      scoreSurface(
          queryTerms,
          List.of(TaskKeywords.inputKindSurface(inputKind)),
          "required input kind",
          8,
          false,
          accumulator);
    }
    for (TaskVerificationKind verificationKind : verificationKinds) {
      scoreSurface(
          queryTerms,
          List.of(TaskKeywords.verificationKindSurface(verificationKind)),
          "verification kind",
          8,
          false,
          accumulator);
    }
  }

  private static void scoreCapabilitySurface(
      List<String> queryTerms, List<TaskPhase> phases, MatchAccumulator accumulator) {
    for (TaskPhase phase : phases) {
      for (TaskCapabilityRef capabilityRef : phase.capabilityRefs()) {
        String capabilityText =
            capabilityRef.id().replace('_', ' ').toLowerCase(java.util.Locale.ROOT);
        scoreSurface(queryTerms, List.of(capabilityText), "capability id", 6, false, accumulator);
        GridGrindProtocolCatalog.entryFor(capabilityRef.qualifiedId())
            .ifPresent(
                entry ->
                    scoreSurface(
                        queryTerms,
                        List.of(entry.summary()),
                        "capability summary",
                        2,
                        false,
                        accumulator));
      }
    }
  }

  private static void scoreSurface(
      List<String> queryTerms,
      List<String> surfaces,
      String surfaceLabel,
      int scorePerTerm,
      boolean semanticSurface,
      MatchAccumulator accumulator) {
    for (String surface : surfaces) {
      List<String> matchedTerms =
          RecipeTerms.intersection(queryTerms, RecipeTerms.normalizedTerms(surface));
      if (matchedTerms.isEmpty()) {
        continue;
      }
      accumulator.score += scorePerTerm * matchedTerms.size();
      accumulator.matchedTerms.addAll(matchedTerms);
      accumulator.hasSemanticMatch |= semanticSurface;
      accumulator.matchSources.add(surfaceLabel);
    }
  }

  private static String exampleStem(String fileName) {
    return fileName.endsWith(".json") ? fileName.substring(0, fileName.length() - 5) : fileName;
  }

  /** Mutable scoring state while one recipe is being matched against one normalized query. */
  private static final class MatchAccumulator {
    private final Set<String> matchedTerms = new LinkedHashSet<>();
    private final Set<String> matchSources = new LinkedHashSet<>();
    private boolean hasSemanticMatch;
    private int score;

    private List<String> compactMatchSources() {
      return matchSources.stream().limit(4).toList();
    }
  }
}
