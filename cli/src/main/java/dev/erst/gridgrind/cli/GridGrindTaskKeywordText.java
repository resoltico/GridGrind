package dev.erst.gridgrind.cli;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Pattern;

/** Normalizes English task-discovery text into deterministic searchable token sets. */
final class GridGrindTaskKeywordText {
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
  private static final List<SingularizationRule> SINGULARIZATION_RULES =
      List.of(
          new SingularizationRule(
              token -> token.length() > 4 && token.endsWith("ies"),
              token -> token.substring(0, token.length() - 3) + "y"),
          new SingularizationRule(
              token -> token.length() > 5 && token.endsWith("zzes"),
              token -> token.substring(0, token.length() - 3)),
          new SingularizationRule(
              token -> token.length() > 4 && token.endsWith("zes"),
              token -> token.substring(0, token.length() - 1)),
          new SingularizationRule(
              token ->
                  token.length() > 4
                      && token.endsWith("es")
                      && (token.endsWith("ches")
                          || token.endsWith("shes")
                          || token.endsWith("xes")
                          || token.endsWith("ses")),
              token -> token.substring(0, token.length() - 2)),
          new SingularizationRule(
              token -> token.length() > 3 && token.endsWith("s") && !token.endsWith("ss"),
              token -> token.substring(0, token.length() - 1)));

  private GridGrindTaskKeywordText() {}

  static List<String> normalizedTerms(String text) {
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

  static List<String> intersection(List<String> queryTerms, List<String> surfaceTerms) {
    Set<String> surface = new LinkedHashSet<>(surfaceTerms);
    List<String> matches = new ArrayList<>();
    for (String queryTerm : queryTerms) {
      if (surface.contains(queryTerm)) {
        matches.add(queryTerm);
      }
    }
    return List.copyOf(matches);
  }

  private static String singularize(String token) {
    for (SingularizationRule rule : SINGULARIZATION_RULES) {
      if (rule.matches().test(token)) {
        return rule.rewrite().apply(token);
      }
    }
    return token;
  }

  private record SingularizationRule(
      java.util.function.Predicate<String> matches,
      java.util.function.UnaryOperator<String> rewrite) {}
}
