package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Pattern;

/** Deterministic task-intent planner layered on top of the contract-owned task catalog. */
public final class GridGrindGoalPlanner {
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
          "of",
          "office",
          "on",
          "or",
          "please",
          "show",
          "sheet",
          "sheets",
          "the",
          "to",
          "want",
          "with",
          "workbook",
          "workbooks",
          "worksheet",
          "worksheets",
          "xlsx");

  private GridGrindGoalPlanner() {}

  /** Returns ranked task matches plus starter scaffolds for one freeform goal string. */
  public static GoalPlanReport reportFor(String goal) {
    String requestedGoal = CatalogRecordValidation.requireNonBlank(goal, "goal");
    List<String> normalizedTerms = normalizedTerms(requestedGoal);
    List<GoalPlanReport.Candidate> candidates =
        GridGrindTaskCatalog.catalog().tasks().stream()
            .map(task -> candidateFor(task, normalizedTerms))
            .filter(Objects::nonNull)
            .sorted(
                java.util.Comparator.comparingInt(GoalPlanReport.Candidate::score)
                    .reversed()
                    .thenComparing(
                        candidate -> candidate.matchedTerms().size(),
                        java.util.Comparator.reverseOrder())
                    .thenComparing(candidate -> candidate.task().id()))
            .toList();
    return new GoalPlanReport(
        GridGrindProtocolVersion.current(),
        requestedGoal,
        normalizedTerms,
        unmatchedTerms(normalizedTerms, candidates),
        suggestedIntentTags(),
        candidates);
  }

  private static GoalPlanReport.Candidate candidateFor(TaskEntry task, List<String> goalTerms) {
    TaskMatchAccumulator accumulator = new TaskMatchAccumulator(task);
    scoreSurface(goalTerms, task.intentTags(), "intent tag", 12, accumulator);
    scoreSurface(goalTerms, List.of(task.summary()), "summary", 5, accumulator);
    scoreSurface(goalTerms, task.outcomes(), "outcome", 4, accumulator);
    scoreSurface(goalTerms, task.requiredInputs(), "required input", 4, accumulator);
    scoreSurface(goalTerms, task.optionalFeatures(), "optional feature", 3, accumulator);
    scorePhaseSurface(goalTerms, task.phases(), accumulator);
    scoreCapabilitySurface(goalTerms, task.phases(), accumulator);
    if (accumulator.score == 0) {
      return null;
    }
    return new GoalPlanReport.Candidate(
        task,
        accumulator.score,
        List.copyOf(accumulator.matchedTerms),
        List.copyOf(accumulator.reasons),
        GridGrindTaskPlanner.planFor(task));
  }

  private static void scorePhaseSurface(
      List<String> goalTerms, List<TaskPhase> phases, TaskMatchAccumulator accumulator) {
    for (TaskPhase phase : phases) {
      scoreSurface(goalTerms, List.of(phase.label()), "phase label", 2, accumulator);
      scoreSurface(goalTerms, List.of(phase.objective()), "phase objective", 2, accumulator);
      scoreSurface(goalTerms, phase.notes(), "phase note", 1, accumulator);
    }
  }

  private static void scoreCapabilitySurface(
      List<String> goalTerms, List<TaskPhase> phases, TaskMatchAccumulator accumulator) {
    for (TaskPhase phase : phases) {
      for (TaskCapabilityRef capabilityRef : phase.capabilityRefs()) {
        String capabilityText = capabilityRef.id().replace('_', ' ').toLowerCase(Locale.ROOT);
        scoreSurface(goalTerms, List.of(capabilityText), "capability", 2, accumulator);
      }
    }
  }

  private static void scoreSurface(
      List<String> goalTerms,
      List<String> surfaces,
      String surfaceLabel,
      int scorePerTerm,
      TaskMatchAccumulator accumulator) {
    for (String surface : surfaces) {
      List<String> matchedTerms = intersection(goalTerms, normalizedTerms(surface));
      if (matchedTerms.isEmpty()) {
        continue;
      }
      accumulator.score += scorePerTerm * matchedTerms.size();
      accumulator.matchedTerms.addAll(matchedTerms);
      accumulator.reasons.add(
          "Matched "
              + surfaceLabel
              + " \""
              + surface
              + "\" via "
              + GridGrindContractText.humanJoin(matchedTerms)
              + ".");
    }
  }

  private static List<String> unmatchedTerms(
      List<String> goalTerms, List<GoalPlanReport.Candidate> candidates) {
    Set<String> matched = new LinkedHashSet<>();
    for (GoalPlanReport.Candidate candidate : candidates) {
      matched.addAll(candidate.matchedTerms());
    }
    return goalTerms.stream().filter(term -> !matched.contains(term)).toList();
  }

  private static List<String> suggestedIntentTags() {
    Set<String> tags = new LinkedHashSet<>();
    for (TaskEntry task : GridGrindTaskCatalog.catalog().tasks()) {
      tags.addAll(task.intentTags());
    }
    return List.copyOf(tags);
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

  private static List<String> intersection(List<String> goalTerms, List<String> surfaceTerms) {
    Set<String> surface = new LinkedHashSet<>(surfaceTerms);
    List<String> matches = new ArrayList<>();
    for (String goalTerm : goalTerms) {
      if (surface.contains(goalTerm)) {
        matches.add(goalTerm);
      }
    }
    return List.copyOf(matches);
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

  /** Mutable scoring state while one task is being matched against one normalized goal. */
  private static final class TaskMatchAccumulator {
    private final Set<String> matchedTerms = new LinkedHashSet<>();
    private final List<String> reasons = new ArrayList<>();
    private int score;

    private TaskMatchAccumulator(TaskEntry task) {
      Objects.requireNonNull(task, "task must not be null");
    }
  }
}
