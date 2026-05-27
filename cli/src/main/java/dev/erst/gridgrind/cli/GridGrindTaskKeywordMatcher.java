package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskCapabilityRef;
import dev.erst.gridgrind.cli.discovery.TaskDiscoveryProfile;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskExecutionProfile;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskKeywordMatchReport;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhase;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Pattern;

/** Deterministic English keyword matcher layered on top of the public task descriptor catalog. */
final class GridGrindTaskKeywordMatcher {
  private static final Pattern TOKEN_SPLIT_PATTERN = Pattern.compile("[^a-z0-9]+");
  private static final Set<String> STOP_WORDS =
      Set.of(
          "a",
          "an",
          "and",
          "as",
          "at",
          "build",
          "create",
          "excel",
          "file",
          "for",
          "from",
          "how",
          "i",
          "in",
          "into",
          "make",
          "my",
          "no",
          "of",
          "office",
          "on",
          "or",
          "please",
          "show",
          "sheet",
          "sheets",
          "such",
          "spreadsheet",
          "spreadsheets",
          "task",
          "tasks",
          "the",
          "to",
          "want",
          "with",
          "workbook",
          "workbooks",
          "workflow",
          "workflows",
          "worksheet",
          "worksheets",
          "xlsx");

  private GridGrindTaskKeywordMatcher() {}

  /** Returns ranked task matches for one English keyword query string. */
  static TaskKeywordMatchReport reportFor(String query) {
    String requestedQuery = Objects.requireNonNull(query, "query must not be null");
    List<String> normalizedTerms = normalizedTerms(requestedQuery);
    if (normalizedTerms.isEmpty()) {
      throw new IllegalArgumentException(
          "query must contain at least one searchable term after normalization");
    }
    List<TaskKeywordMatchReport.Candidate> candidates =
        GridGrindTaskCatalog.catalog().tasks().stream()
            .map(task -> candidateFor(task, normalizedTerms))
            .flatMap(Optional::stream)
            .sorted(candidateOrdering())
            .toList();
    return new TaskKeywordMatchReport(
        GridGrindProtocolVersion.current(),
        requestedQuery,
        normalizedTerms,
        unmatchedTerms(normalizedTerms, candidates),
        suggestedIntentTags(normalizedTerms, candidates),
        candidates);
  }

  private static Optional<TaskKeywordMatchReport.Candidate> candidateFor(
      TaskEntry task, List<String> queryTerms) {
    TaskMatchAccumulator accumulator = new TaskMatchAccumulator(task);
    scoreSurface(
        queryTerms, List.of(task.id().replace('_', ' ')), "task id", 14, true, accumulator);
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
        new TaskKeywordMatchReport.Candidate(
            task.id(),
            task.narrative().summary(),
            accumulator.score,
            List.copyOf(accumulator.matchedTerms),
            List.copyOf(accumulator.matchSources)));
  }

  private static void scoreDiscoveryProfile(
      List<String> queryTerms,
      TaskDiscoveryProfile discoveryProfile,
      TaskMatchAccumulator accumulator) {
    scoreSurface(
        queryTerms, discoveryProfile.discoveryTerms(), "discovery term", 13, true, accumulator);
    scoreSurface(queryTerms, discoveryProfile.intentTags(), "intent tag", 11, true, accumulator);
  }

  private static void scorePhaseSurface(
      List<String> queryTerms, List<TaskPhase> phases, TaskMatchAccumulator accumulator) {
    for (TaskPhase phase : phases) {
      scoreSurface(
          queryTerms,
          List.of(phasePurposeSurface(phase.purpose())),
          "phase purpose",
          4,
          true,
          accumulator);
      scoreSurface(queryTerms, List.of(phase.label()), "phase label", 2, true, accumulator);
      scoreSurface(
          queryTerms, List.of(phase.objective()), "phase objective", 1, false, accumulator);
    }
  }

  private static void scoreExecutionProfileSurface(
      List<String> queryTerms, TaskExecutionProfile profile, TaskMatchAccumulator accumulator) {
    scoreSurface(
        queryTerms,
        List.of(sourceModeSurface(profile.sourceMode())),
        "source mode",
        6,
        false,
        accumulator);
    scoreSurface(
        queryTerms,
        List.of(persistenceModeSurface(profile.persistenceMode())),
        "persistence mode",
        6,
        false,
        accumulator);
    scoreSurface(
        queryTerms,
        List.of(mutationModeSurface(profile.mutationMode())),
        "mutation mode",
        6,
        false,
        accumulator);
    scoreSurface(
        queryTerms,
        List.of(assetModeSurface(profile.assetMode())),
        "asset mode",
        5,
        false,
        accumulator);
  }

  private static void scoreTypedTaskSurface(
      List<String> queryTerms,
      List<TaskGoalKind> goals,
      List<TaskArtifactKind> artifacts,
      List<TaskInputKind> requiredInputKinds,
      List<TaskVerificationKind> verificationKinds,
      TaskMatchAccumulator accumulator) {
    for (TaskGoalKind goal : goals) {
      scoreSurface(queryTerms, List.of(goalSurface(goal)), "task goal", 10, true, accumulator);
    }
    for (TaskArtifactKind artifact : artifacts) {
      scoreSurface(
          queryTerms, List.of(artifactSurface(artifact)), "artifact", 3, false, accumulator);
    }
    for (TaskInputKind inputKind : requiredInputKinds) {
      scoreSurface(
          queryTerms,
          List.of(inputKindSurface(inputKind)),
          "required input kind",
          8,
          false,
          accumulator);
    }
    for (TaskVerificationKind verificationKind : verificationKinds) {
      scoreSurface(
          queryTerms,
          List.of(verificationKindSurface(verificationKind)),
          "verification kind",
          8,
          false,
          accumulator);
    }
  }

  private static void scoreCapabilitySurface(
      List<String> queryTerms, List<TaskPhase> phases, TaskMatchAccumulator accumulator) {
    for (TaskPhase phase : phases) {
      for (TaskCapabilityRef capabilityRef : phase.capabilityRefs()) {
        String capabilityText = capabilityRef.id().replace('_', ' ').toLowerCase(Locale.ROOT);
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
      TaskMatchAccumulator accumulator) {
    for (String surface : surfaces) {
      List<String> matchedTerms = intersection(queryTerms, normalizedTerms(surface));
      if (matchedTerms.isEmpty()) {
        continue;
      }
      accumulator.score += scorePerTerm * matchedTerms.size();
      accumulator.matchedTerms.addAll(matchedTerms);
      accumulator.hasSemanticMatch |= semanticSurface;
      accumulator.matchSources.add(surfaceLabel);
    }
  }

  private static List<String> unmatchedTerms(
      List<String> queryTerms, List<TaskKeywordMatchReport.Candidate> candidates) {
    Set<String> matched = new LinkedHashSet<>();
    for (TaskKeywordMatchReport.Candidate candidate : candidates) {
      matched.addAll(candidate.matchedTerms());
    }
    return queryTerms.stream().filter(term -> !matched.contains(term)).toList();
  }

  static Comparator<TaskKeywordMatchReport.Candidate> candidateOrdering() {
    return Comparator.comparingInt(TaskKeywordMatchReport.Candidate::score)
        .reversed()
        .thenComparing(candidate -> candidate.matchedTerms().size(), Comparator.reverseOrder())
        .thenComparing(TaskKeywordMatchReport.Candidate::taskId);
  }

  static List<String> suggestedIntentTags(
      List<String> queryTerms, List<TaskKeywordMatchReport.Candidate> candidates) {
    if (candidates.isEmpty()) {
      return List.of();
    }
    int topScore = candidates.getFirst().score();
    List<TagSuggestion> suggestionScores = new ArrayList<>();
    for (TaskKeywordMatchReport.Candidate candidate : candidates) {
      if (candidate.score() * 4 < topScore) {
        continue;
      }
      TaskEntry task = GridGrindTaskCatalog.entryFor(candidate.taskId()).orElseThrow();
      for (String tag : task.discoveryProfile().intentTags()) {
        int overlapCount = intersection(queryTerms, normalizedTerms(tag)).size();
        int weightedScore = candidate.score() + (overlapCount * 10_000);
        mergeTagSuggestion(suggestionScores, tag, weightedScore);
      }
    }
    return suggestionScores.stream()
        .sorted(
            Comparator.comparingInt(TagSuggestion::score)
                .reversed()
                .thenComparing(TagSuggestion::tag))
        .limit(8)
        .map(TagSuggestion::tag)
        .toList();
  }

  private static void mergeTagSuggestion(
      List<TagSuggestion> suggestions, String tag, int weightedScore) {
    int existingIndex = -1;
    for (int index = 0; index < suggestions.size(); index++) {
      if (suggestions.get(index).tag().equals(tag)) {
        existingIndex = index;
        break;
      }
    }
    if (existingIndex < 0) {
      suggestions.add(new TagSuggestion(tag, weightedScore));
      return;
    }
    TagSuggestion existing = suggestions.get(existingIndex);
    if (weightedScore > existing.score()) {
      suggestions.set(existingIndex, new TagSuggestion(tag, weightedScore));
    }
  }

  private static List<String> normalizedTerms(String text) {
    Set<String> normalized = new LinkedHashSet<>();
    for (String rawToken :
        TOKEN_SPLIT_PATTERN
            .splitAsStream(text.toLowerCase(Locale.ROOT).replace('_', ' '))
            .toList()) {
      if (rawToken.isBlank() || STOP_WORDS.contains(rawToken)) {
        continue;
      }
      normalized.add(singularize(rawToken));
    }
    return List.copyOf(normalized);
  }

  private static List<String> intersection(List<String> queryTerms, List<String> surfaceTerms) {
    Set<String> surface = new LinkedHashSet<>(surfaceTerms);
    List<String> matches = new ArrayList<>();
    for (String queryTerm : queryTerms) {
      if (surface.contains(queryTerm)) {
        matches.add(queryTerm);
      }
    }
    return List.copyOf(matches);
  }

  static String sourceModeSurface(TaskSourceMode sourceMode) {
    return switch (sourceMode) {
      case NEW_WORKBOOK -> "new blank create workbook";
      case EXISTING_WORKBOOK -> "existing open input workbook";
    };
  }

  static String persistenceModeSurface(TaskPersistenceMode persistenceMode) {
    return switch (persistenceMode) {
      case NONE -> "no save transient non destructive";
      case SAVE_AS -> "save export output";
      case OVERWRITE_SOURCE -> "overwrite in place";
    };
  }

  static String mutationModeSurface(TaskMutationMode mutationMode) {
    return switch (mutationMode) {
      case READ_ONLY -> "read only inspect audit analyze";
      case MUTATING -> "author mutate build update";
    };
  }

  static String assetModeSurface(TaskAssetMode assetMode) {
    return switch (assetMode) {
      case SELF_CONTAINED -> "self contained portable";
      case REQUIRES_EXTERNAL_PAYLOADS -> "external payload file assets";
    };
  }

  static String phasePurposeSurface(TaskPhasePurpose purpose) {
    return switch (purpose) {
      case PREPARE -> "prepare setup scaffold";
      case AUTHOR -> "author build mutate";
      case IMPORT -> "import ingest";
      case EXPORT -> "export extract";
      case INSPECT -> "inspect readback";
      case ANALYZE -> "analyze findings health";
      case VERIFY -> "verify assert confirm";
    };
  }

  static String inputKindSurface(TaskInputKind inputKind) {
    return switch (inputKind) {
      case SOURCE_WORKBOOK_PATH -> "existing workbook input source file path";
      case PERSISTENCE_TARGET_PATH -> "save export output destination path";
      case TARGET_SHEET_NAMES -> "sheet name worksheet names";
      case TARGET_OBJECT_NAMES -> "table pivot chart named range drawing object names";
      case CELL_OR_RANGE_COORDINATES -> "cell range coordinates address anchor placement";
      case TABULAR_SOURCE_ROWS -> "source rows table data values";
      case VALIDATION_RULES -> "validation rules allowed values protection prompts";
      case MAPPING_LOCATOR -> "mapping locator map id mapping name";
      case XML_PAYLOAD -> "xml payload mapping document import export";
      case BINARY_PAYLOAD -> "binary payload image object file";
      case DRAWING_ANCHORS -> "drawing anchor position placement";
    };
  }

  static String goalSurface(TaskGoalKind goal) {
    return switch (goal) {
      case AUTHOR -> "author create build";
      case VERIFY -> "verify validate confirm";
      case INSPECT -> "inspect read discover";
      case ANALYZE -> "analyze findings diagnose";
      case IMPORT -> "import ingest load";
      case EXPORT -> "export extract emit";
      case MAINTAIN -> "maintain normalize repair";
    };
  }

  static String artifactSurface(TaskArtifactKind artifact) {
    return switch (artifact) {
      case WORKBOOK -> "workbook spreadsheet file";
      case SHEET -> "sheet worksheet tab";
      case CELL -> "cell grid value";
      case TABLE -> "table rows columns";
      case CHART -> "chart graph dashboard";
      case PIVOT_TABLE -> "pivot pivot table summary";
      case DATA_VALIDATION -> "validation input rules";
      case COMMENT -> "comment annotation remark";
      case PROTECTION -> "protection locked guardrail";
      case CUSTOM_XML_MAPPING -> "custom xml mapping";
      case XML_PAYLOAD -> "xml payload import export";
      case DRAWING_OBJECT -> "drawing picture shape object";
      case SIGNATURE_LINE -> "signature line signing";
      case PACKAGE_SECURITY -> "package security encryption ooxml";
      case FORMULA_SURFACE -> "formula surface formulas";
      case NAMED_RANGE -> "named range name reference";
    };
  }

  static String verificationKindSurface(TaskVerificationKind verificationKind) {
    return switch (verificationKind) {
      case FACT_READBACK -> "inspect readback factual verification";
      case ASSERTION_CHECKS -> "assert assertion invariant checks";
      case HEALTH_ANALYSIS -> "health analysis findings audit";
      case EXPORT_REREAD -> "export reread roundtrip verification";
    };
  }

  private static String singularize(String token) {
    if (token.length() > 4 && token.endsWith("ies")) {
      return token.substring(0, token.length() - 3) + "y";
    }
    if (token.length() > 5 && token.endsWith("zzes")) {
      return token.substring(0, token.length() - 3);
    }
    if (token.length() > 4 && token.endsWith("zes")) {
      return token.substring(0, token.length() - 1);
    }
    if (token.length() > 4
        && token.endsWith("es")
        && (token.endsWith("ches")
            || token.endsWith("shes")
            || token.endsWith("xes")
            || token.endsWith("ses"))) {
      return token.substring(0, token.length() - 2);
    }
    if (token.length() > 3 && token.endsWith("s") && !token.endsWith("ss")) {
      return token.substring(0, token.length() - 1);
    }
    return token;
  }

  /** Mutable scoring state while one task is being matched against one normalized query. */
  private static final class TaskMatchAccumulator {
    private final Set<String> matchedTerms = new LinkedHashSet<>();
    private final Set<String> matchSources = new LinkedHashSet<>();
    private boolean hasSemanticMatch;
    private int score;

    private TaskMatchAccumulator(TaskEntry task) {
      Objects.requireNonNull(task, "task must not be null");
    }
  }

  private record TagSuggestion(String tag, int score) {}
}
