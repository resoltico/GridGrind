package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindContractVocabulary;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.junit.jupiter.api.Test;

/** Build-failing lint over CLI-owned public help text that mentions contract ids. */
class GridGrindCliPublicSurfaceLintTest {
  @Test
  void helpTextReferencesOnlyRegisteredCatalogIds() {
    Set<String> registeredIds = GridGrindContractVocabulary.registeredPublicSurfaceIds();
    Pattern candidatePattern = GridGrindContractVocabulary.candidateIdPattern();
    String helpText =
        GridGrindCliHelp.helpText(
            "dev", "CLI lint surface", "https://example.invalid/gridgrind", "gridgrind:test");

    Set<String> unknown = collectUnknown(helpText, registeredIds, candidatePattern);

    assertTrue(
        unknown.isEmpty(),
        () -> "Unregistered public step ids leaked into CLI help text: " + unknown);
  }

  private static Set<String> collectUnknown(
      String text, Set<String> registeredIds, Pattern candidatePattern) {
    Set<String> unknown = new LinkedHashSet<>();
    Matcher matcher = candidatePattern.matcher(text);
    while (matcher.find()) {
      String candidate = matcher.group();
      if (registeredIds.contains(candidate)) {
        continue;
      }
      String candidatePrefix = prefixOf(candidate);
      boolean sharesKnownPrefix =
          registeredIds.stream()
              .map(GridGrindCliPublicSurfaceLintTest::prefixOf)
              .anyMatch(candidatePrefix::equals);
      if (sharesKnownPrefix && (candidate.contains("*") || candidate.endsWith("_"))) {
        unknown.add(candidate);
      }
    }
    return unknown;
  }

  private static String prefixOf(String candidate) {
    int separator = candidate.indexOf('_');
    return separator < 0 ? candidate : candidate.substring(0, separator);
  }
}
