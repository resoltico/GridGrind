package dev.erst.gridgrind.buildlogic;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;
import org.gradle.api.GradleException;

public final class JavaSemanticShapePolicy {
  private final List<Rule> rules;

  private JavaSemanticShapePolicy(List<Rule> rules) {
    this.rules = List.copyOf(rules);
  }

  static JavaSemanticShapePolicy load(Path policyFile) throws IOException {
    List<Rule> loadedRules = new ArrayList<>();
    int lineNumber = 0;
    for (String rawLine : Files.readAllLines(policyFile)) {
      lineNumber++;
      String line = rawLine.trim();
      if (line.isEmpty() || line.startsWith("#")) {
        continue;
      }
      String[] columns = rawLine.split("\t", -1);
      if (columns.length != 8) {
        throw new GradleException(
            "Semantic-shape policy lines must contain 8 tab-separated columns, but line "
                + lineNumber
                + " in "
                + policyFile
                + " had "
                + columns.length
                + ".");
      }
      loadedRules.add(
          new Rule(
              loadedRules.size(),
              parseMatchKind(columns[0], lineNumber, policyFile),
              normalizeRulePath(columns[1], lineNumber, policyFile),
              requiredValue("role", columns[2], lineNumber, policyFile),
              parseAllowedRules(columns[3], lineNumber, policyFile),
              requiredValue("owner", columns[4], lineNumber, policyFile),
              parseDate("reviewExpiresOn", columns[5], lineNumber, policyFile),
              requiredValue("splitTrigger", columns[6], lineNumber, policyFile),
              requiredValue("rationale", columns[7], lineNumber, policyFile)));
    }
    validateRules(loadedRules, policyFile);
    return new JavaSemanticShapePolicy(loadedRules);
  }

  List<Rule> rules() {
    return rules;
  }

  Rule matchingRule(String relativePath) {
    return rules.stream()
        .filter(rule -> rule.matches(relativePath))
        .sorted(
            java.util.Comparator.comparingInt(JavaSemanticShapePolicy::matchPriority)
                .thenComparing(
                    java.util.Comparator.comparingInt(
                            (Rule rule) ->
                                rule.kind() == JavaSourceShapePolicy.MatchKind.PREFIX
                                    ? rule.path().length()
                                    : 0)
                        .reversed())
                .thenComparingInt(Rule::order))
        .findFirst()
        .orElse(null);
  }

  private static int matchPriority(Rule rule) {
    return switch (rule.kind()) {
      case EXACT -> 0;
      case PREFIX -> 1;
      case DEFAULT -> 2;
    };
  }

  private static void validateRules(List<Rule> rules, Path policyFile) {
    if (rules.isEmpty()) {
      throw new GradleException(
          "Semantic-shape policy " + policyFile + " must declare at least one rule.");
    }
    if (rules.stream().anyMatch(rule -> rule.kind() == JavaSourceShapePolicy.MatchKind.DEFAULT)) {
      throw new GradleException(
          "Semantic-shape policy " + policyFile + " must not declare DEFAULT rules.");
    }
    Set<String> exactPaths = new LinkedHashSet<>();
    for (Rule rule : rules) {
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.EXACT && !exactPaths.add(rule.path())) {
        throw new GradleException(
            "Semantic-shape policy "
                + policyFile
                + " declares duplicate EXACT rules for "
                + rule.path()
                + ".");
      }
    }
  }

  private static String requiredValue(
      String fieldName, String rawValue, int lineNumber, Path policyFile) {
    String value = rawValue.trim();
    if (value.isEmpty()) {
      throw new GradleException(
          "Semantic-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " must declare "
              + fieldName
              + ".");
    }
    return value;
  }

  private static Set<String> parseAllowedRules(
      String rawValue, int lineNumber, Path policyFile) {
    String value = rawValue.trim();
    if (value.isEmpty()) {
      throw new GradleException(
          "Semantic-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " must declare allowedRules.");
    }
    Set<String> rules =
        java.util.Arrays.stream(value.split(","))
            .map(String::trim)
            .filter(ruleName -> !ruleName.isEmpty())
            .collect(Collectors.toCollection(LinkedHashSet::new));
    if (rules.isEmpty()) {
      throw new GradleException(
          "Semantic-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " must declare at least one allowed rule.");
    }
    return Set.copyOf(rules);
  }

  private static LocalDate parseDate(
      String fieldName, String rawValue, int lineNumber, Path policyFile) {
    String value = rawValue.trim();
    if (value.isEmpty() || "-".equals(value)) {
      throw new GradleException(
          "Semantic-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " must declare "
              + fieldName
              + ".");
    }
    try {
      return LocalDate.parse(value);
    } catch (RuntimeException exception) {
      throw new GradleException(
          "Semantic-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " has invalid "
              + fieldName
              + " value '"
              + rawValue
              + "'.",
          exception);
    }
  }

  private static String normalizeRulePath(String rawValue, int lineNumber, Path policyFile) {
    String normalized = rawValue.trim().replace('\\', '/');
    if (normalized.startsWith("./")) {
      normalized = normalized.substring(2);
    }
    if (normalized.startsWith("/")) {
      throw new GradleException(
          "Semantic-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " must use repository-relative paths.");
    }
    return normalized;
  }

  private static JavaSourceShapePolicy.MatchKind parseMatchKind(
      String rawValue, int lineNumber, Path policyFile) {
    try {
      return JavaSourceShapePolicy.MatchKind.valueOf(rawValue.trim());
    } catch (IllegalArgumentException exception) {
      throw new GradleException(
          "Semantic-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " uses unknown rule kind '"
              + rawValue
              + "'.",
          exception);
    }
  }

  record Rule(
      int order,
      JavaSourceShapePolicy.MatchKind kind,
      String path,
      String role,
      Set<String> allowedRules,
      String owner,
      LocalDate reviewExpiresOn,
      String splitTrigger,
      String rationale) {
    Rule {
      if (kind == JavaSourceShapePolicy.MatchKind.DEFAULT) {
        throw new GradleException("Semantic-shape rules must not use DEFAULT.");
      }
      if (path.isEmpty()) {
        throw new GradleException("Semantic-shape rules must declare a path.");
      }
      if (allowedRules.isEmpty()) {
        throw new GradleException("Semantic-shape rules must declare allowedRules.");
      }
    }

    boolean matches(String relativePath) {
      return switch (kind) {
        case EXACT -> path.equals(relativePath);
        case PREFIX -> relativePath.startsWith(path);
        case DEFAULT -> false;
      };
    }

    boolean allows(String ruleName) {
      return allowedRules.contains(ruleName);
    }
  }
}
