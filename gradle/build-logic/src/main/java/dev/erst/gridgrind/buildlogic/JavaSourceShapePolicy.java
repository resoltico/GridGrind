package dev.erst.gridgrind.buildlogic;

import java.io.IOException;
import java.time.LocalDate;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import org.gradle.api.GradleException;

public final class JavaSourceShapePolicy {
  private final List<Rule> rules;

  private JavaSourceShapePolicy(List<Rule> rules) {
    this.rules = List.copyOf(rules);
  }

  static JavaSourceShapePolicy load(Path policyFile) throws IOException {
    List<Rule> loadedRules = new ArrayList<>();
    int lineNumber = 0;
    for (String rawLine : Files.readAllLines(policyFile)) {
      lineNumber++;
      String line = rawLine.trim();
      if (line.isEmpty() || line.startsWith("#")) {
        continue;
      }
      String[] columns = rawLine.split("\t", -1);
      if (columns.length != 15) {
        throw new GradleException(
            "Source-shape policy lines must contain 15 tab-separated columns, but line "
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
              MatchKind.parse(columns[0], lineNumber, policyFile),
              normalizeRulePath(columns[1], lineNumber, policyFile),
              requiredValue("role", columns[2], lineNumber, policyFile),
              parseLimit("maxLines", columns[3], lineNumber, policyFile),
              parseLimit("maxMethods", columns[4], lineNumber, policyFile),
              parseLimit("maxPublicMethods", columns[5], lineNumber, policyFile),
              parseLimit("maxImports", columns[6], lineNumber, policyFile),
              parseLimit("maxFields", columns[7], lineNumber, policyFile),
              parseLimit("maxSwitches", columns[8], lineNumber, policyFile),
              parseLimit("maxSwitchArms", columns[9], lineNumber, policyFile),
              requiredValue("owner", columns[10], lineNumber, policyFile),
              DuplicationGuard.parse(columns[11], lineNumber, policyFile),
              parseDate("reviewExpiresOn", columns[12], lineNumber, policyFile),
              optionalValue(columns[13]),
              requiredValue("rationale", columns[14], lineNumber, policyFile)));
    }
    validateRules(loadedRules, policyFile);
    return new JavaSourceShapePolicy(loadedRules);
  }

  List<Rule> rules() {
    return rules;
  }

  List<Rule> matchingRules(String relativePath) {
    return rules.stream()
        .filter(rule -> rule.matches(relativePath))
        .sorted(
            java.util.Comparator.comparingInt(JavaSourceShapePolicy::matchPriority)
                .thenComparing(
                    java.util.Comparator.comparingInt(
                            (Rule rule) -> rule.kind() == MatchKind.PREFIX ? rule.path().length() : 0)
                        .reversed())
                .thenComparingInt(Rule::order))
        .toList();
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
      throw new GradleException("Source-shape policy " + policyFile + " must declare at least one rule.");
    }

    long defaultRuleCount = rules.stream().filter(rule -> rule.kind() == MatchKind.DEFAULT).count();
    if (defaultRuleCount != 1L) {
      throw new GradleException(
          "Source-shape policy "
              + policyFile
              + " must declare exactly one DEFAULT rule, but declared "
              + defaultRuleCount
              + ".");
    }
    if (rules.getLast().kind() != MatchKind.DEFAULT) {
      throw new GradleException(
          "Source-shape policy " + policyFile + " must place the DEFAULT rule last.");
    }

    Set<String> exactPaths = new LinkedHashSet<>();
    for (Rule rule : rules) {
      if (rule.kind() == MatchKind.EXACT && !exactPaths.add(rule.path())) {
        throw new GradleException(
            "Source-shape policy "
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
          "Source-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " must declare "
              + fieldName
              + ".");
    }
    return value;
  }

  private static Integer parseLimit(
      String fieldName, String rawValue, int lineNumber, Path policyFile) {
    String value = rawValue.trim();
    if (value.isEmpty() || "-".equals(value)) {
      return null;
    }
    try {
      int parsed = Integer.parseInt(value);
      if (parsed < 0) {
        throw new NumberFormatException("negative");
      }
      return parsed;
    } catch (NumberFormatException exception) {
      throw new GradleException(
          "Source-shape policy "
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

  private static LocalDate parseDate(
      String fieldName, String rawValue, int lineNumber, Path policyFile) {
    String value = rawValue.trim();
    if (value.isEmpty() || "-".equals(value)) {
      return null;
    }
    try {
      return LocalDate.parse(value);
    } catch (RuntimeException exception) {
      throw new GradleException(
          "Source-shape policy "
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

  private static String optionalValue(String rawValue) {
    String value = rawValue.trim();
    return value.isEmpty() || "-".equals(value) ? null : value;
  }

  private static String normalizeRulePath(String rawValue, int lineNumber, Path policyFile) {
    String normalized = rawValue.trim().replace('\\', '/');
    if (normalized.startsWith("./")) {
      normalized = normalized.substring(2);
    }
    if (normalized.startsWith("/")) {
      throw new GradleException(
          "Source-shape policy "
              + policyFile
              + " line "
              + lineNumber
              + " must use repository-relative paths.");
    }
    return normalized;
  }

  enum MatchKind {
    EXACT,
    PREFIX,
    DEFAULT;

    private static MatchKind parse(String rawValue, int lineNumber, Path policyFile) {
      try {
        return MatchKind.valueOf(rawValue.trim());
      } catch (IllegalArgumentException exception) {
        throw new GradleException(
            "Source-shape policy "
                + policyFile
                + " line "
                + lineNumber
                + " uses unknown rule kind '"
                + rawValue
                + "'.",
            exception);
      }
    }
  }

  enum DuplicationGuard {
    CHECK,
    SKIP;

    private static DuplicationGuard parse(String rawValue, int lineNumber, Path policyFile) {
      try {
        return DuplicationGuard.valueOf(rawValue.trim());
      } catch (IllegalArgumentException exception) {
        throw new GradleException(
            "Source-shape policy "
                + policyFile
                + " line "
                + lineNumber
                + " uses unknown duplicationGuard '"
                + rawValue
                + "'.",
            exception);
      }
    }
  }

  record Rule(
      int order,
      MatchKind kind,
      String path,
      String role,
      Integer maxLines,
      Integer maxMethods,
      Integer maxPublicMethods,
      Integer maxImports,
      Integer maxFields,
      Integer maxSwitches,
      Integer maxSwitchArms,
      String owner,
      DuplicationGuard duplicationGuard,
      LocalDate reviewExpiresOn,
      String splitTrigger,
      String rationale) {
    Rule {
      if (kind != MatchKind.DEFAULT && path.isEmpty()) {
        throw new GradleException("Non-default source-shape rules must declare a path.");
      }
      if (kind == MatchKind.DEFAULT && !("*".equals(path) || path.isEmpty())) {
        throw new GradleException("DEFAULT source-shape rules must use '*' as their path.");
      }
      if (duplicationGuard == null) {
        throw new GradleException("Source-shape rules must declare duplicationGuard.");
      }
      if (kind == MatchKind.EXACT) {
        if (reviewExpiresOn == null) {
          throw new GradleException("EXACT source-shape rules must declare reviewExpiresOn.");
        }
        if (splitTrigger == null || splitTrigger.isBlank()) {
          throw new GradleException("EXACT source-shape rules must declare splitTrigger.");
        }
      } else {
        if (reviewExpiresOn != null) {
          throw new GradleException(
              "Only EXACT source-shape rules may declare reviewExpiresOn.");
        }
        if (splitTrigger != null) {
          throw new GradleException("Only EXACT source-shape rules may declare splitTrigger.");
        }
      }
    }

    boolean matches(String relativePath) {
      return switch (kind) {
        case EXACT -> path.equals(relativePath);
        case PREFIX -> relativePath.startsWith(path);
        case DEFAULT -> true;
      };
    }
  }
}
