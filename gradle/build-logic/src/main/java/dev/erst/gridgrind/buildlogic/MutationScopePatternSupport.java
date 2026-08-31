package dev.erst.gridgrind.buildlogic;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/** Matches the PIT wildcard vocabulary against compiled class names deterministically. */
final class MutationScopePatternSupport {
  private MutationScopePatternSupport() {}

  static List<String> matchingClassNames(Collection<String> classNames, String pattern) {
    return classNames.stream().filter(className -> matches(className, pattern)).sorted().toList();
  }

  static boolean matches(String className, String pattern) {
    List<String> literalParts = literalParts(pattern);
    int searchFrom = 0;
    for (int index = 0; index < literalParts.size(); index++) {
      String literalPart = literalParts.get(index);
      int foundAt = className.indexOf(literalPart, searchFrom);
      if (foundAt < 0 || (index == 0 && !pattern.startsWith("*") && foundAt != 0)) {
        return false;
      }
      searchFrom = foundAt + literalPart.length();
    }
    return pattern.endsWith("*") || searchFrom == className.length();
  }

  private static List<String> literalParts(String pattern) {
    List<String> literalParts = new ArrayList<>();
    StringBuilder literal = new StringBuilder();
    for (int index = 0; index < pattern.length(); index++) {
      char character = pattern.charAt(index);
      if (character == '*') {
        if (!literal.isEmpty()) {
          literalParts.add(literal.toString());
          literal.setLength(0);
        }
        continue;
      }
      literal.append(character);
    }
    if (!literal.isEmpty()) {
      literalParts.add(literal.toString());
    }
    return literalParts;
  }
}
