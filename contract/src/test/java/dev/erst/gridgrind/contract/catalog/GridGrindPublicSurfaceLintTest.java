package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Build-failing lint over public-facing surfaces that mention canonical contract step ids. */
class GridGrindPublicSurfaceLintTest {
  @Test
  void publicSurfacesReferenceOnlyRegisteredCatalogIds() throws IOException {
    Set<String> registeredIds = GridGrindContractVocabulary.registeredPublicSurfaceIds();
    Pattern candidatePattern = GridGrindContractVocabulary.candidateIdPattern();
    Map<String, Set<String>> unknownBySurface = new ConcurrentHashMap<>();

    collectUnknown(
        unknownBySurface,
        "CLI help",
        GridGrindCliHelp.helpText(
            "dev", "Contract lint surface", "https://example.invalid/gridgrind", "gridgrind:test"),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "Protocol catalog summaries",
        catalogSummaries(),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "Request document limit message",
        GridGrindContractText.requestDocumentTooLargeMessage(),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "STANDARD_INPUT message",
        GridGrindContractText.standardInputRequiresRequestMessage(),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "EVENT_READ calculation failure",
        GridGrindExecutionModeMetadata.eventRead().calculationFailureMessage(),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "EVENT_READ unsupported step",
        GridGrindExecutionModeMetadata.eventRead().unsupportedStepMessage("MUTATION"),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "EVENT_READ unsupported query",
        GridGrindExecutionModeMetadata.eventRead().unsupportedQueryMessage("GET_WINDOW"),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "STREAMING_WRITE calculation failure",
        GridGrindExecutionModeMetadata.streamingWrite().calculationFailureMessage(),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "STREAMING_WRITE source failure",
        GridGrindExecutionModeMetadata.streamingWrite().invalidSourceMessage(),
        registeredIds,
        candidatePattern);
    collectUnknown(
        unknownBySurface,
        "STREAMING_WRITE unsupported action",
        GridGrindExecutionModeMetadata.streamingWrite().unsupportedActionMessage("SET_CELL"),
        registeredIds,
        candidatePattern);

    Path repositoryRoot = repositoryRoot();
    collectUnknown(
        unknownBySurface, repositoryRoot.resolve("README.md"), registeredIds, candidatePattern);
    try (Stream<Path> docs = Files.walk(repositoryRoot.resolve("docs"));
        Stream<Path> examples = Files.walk(repositoryRoot.resolve("examples"))) {
      docs.filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".md"))
          .forEach(path -> collectUnknown(unknownBySurface, path, registeredIds, candidatePattern));
      examples
          .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".json"))
          .forEach(path -> collectUnknown(unknownBySurface, path, registeredIds, candidatePattern));
    }

    assertTrue(
        unknownBySurface.isEmpty(),
        () -> "Unregistered public step ids leaked into public surfaces: " + unknownBySurface);
  }

  @Test
  void candidatePatternStillFlagsBrokenContractTokenShapes() {
    Pattern candidatePattern = GridGrindContractVocabulary.candidateIdPattern();
    Set<String> unknown =
        collectUnknown(
            "DELETE_ and BOGUS_PROTOCOL_TOKEN",
            GridGrindContractVocabulary.registeredPublicSurfaceIds(),
            candidatePattern);

    assertEquals(Set.of("DELETE_"), unknown);
  }

  @Test
  void candidatePatternIgnoresUnrelatedUppercaseVocabulary() {
    Pattern candidatePattern = GridGrindContractVocabulary.candidateIdPattern();
    Set<String> unknown =
        collectUnknown(
            "`MOVE_AND_RESIZE` and `UNSUPPORTED_FORMULA`",
            GridGrindContractVocabulary.registeredPublicSurfaceIds(),
            candidatePattern);

    assertTrue(unknown.isEmpty(), () -> "unexpected lint matches: " + unknown);
  }

  private static void collectUnknown(
      Map<String, Set<String>> unknownBySurface,
      Path path,
      java.util.Set<String> registeredIds,
      Pattern candidatePattern) {
    try {
      collectUnknown(
          unknownBySurface,
          path.toString(),
          Files.readString(path),
          registeredIds,
          candidatePattern);
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to read lint surface " + path, exception);
    }
  }

  private static void collectUnknown(
      Map<String, Set<String>> unknownBySurface,
      String surfaceName,
      String text,
      java.util.Set<String> registeredIds,
      Pattern candidatePattern) {
    Set<String> unknown = collectUnknown(text, registeredIds, candidatePattern);
    if (!unknown.isEmpty()) {
      unknownBySurface.put(surfaceName, unknown);
    }
  }

  private static Set<String> collectUnknown(
      String text, java.util.Set<String> registeredIds, Pattern candidatePattern) {
    Set<String> unknown = new LinkedHashSet<>();
    Matcher matcher = candidatePattern.matcher(text);
    while (matcher.find()) {
      String candidate = matcher.group();
      if (isSuspiciousUnknownCandidate(candidate, registeredIds)) {
        unknown.add(candidate);
      }
    }
    return unknown;
  }

  private static boolean isSuspiciousUnknownCandidate(
      String candidate, java.util.Set<String> registeredIds) {
    if (registeredIds.contains(candidate)) {
      return false;
    }
    String candidatePrefix = prefixOf(candidate);
    List<String> samePrefixIds =
        registeredIds.stream().filter(id -> prefixOf(id).equals(candidatePrefix)).toList();
    if (samePrefixIds.isEmpty()) {
      return false;
    }
    if (candidate.contains("*") || candidate.endsWith("_")) {
      return true;
    }
    return samePrefixIds.stream().anyMatch(id -> levenshteinDistance(id, candidate) <= 2);
  }

  private static int levenshteinDistance(String left, String right) {
    int[] previous = new int[right.length() + 1];
    int[] current = new int[right.length() + 1];
    for (int column = 0; column <= right.length(); column++) {
      previous[column] = column;
    }
    for (int row = 1; row <= left.length(); row++) {
      current[0] = row;
      for (int column = 1; column <= right.length(); column++) {
        int substitutionCost = left.charAt(row - 1) == right.charAt(column - 1) ? 0 : 1;
        current[column] =
            Math.min(
                Math.min(current[column - 1] + 1, previous[column] + 1),
                previous[column - 1] + substitutionCost);
      }
      int[] swap = previous;
      previous = current;
      current = swap;
    }
    return previous[right.length()];
  }

  private static String prefixOf(String id) {
    int underscore = id.indexOf('_');
    return underscore < 0 ? id : id.substring(0, underscore);
  }

  private static String catalogSummaries() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    List<String> summaries = new ArrayList<>();
    summaries.add(catalog.requestType().summary());
    catalog.sourceTypes().stream().map(TypeEntry::summary).forEach(summaries::add);
    catalog.persistenceTypes().stream().map(TypeEntry::summary).forEach(summaries::add);
    catalog.stepTypes().stream().map(TypeEntry::summary).forEach(summaries::add);
    catalog.mutationActionTypes().stream().map(TypeEntry::summary).forEach(summaries::add);
    catalog.assertionTypes().stream().map(TypeEntry::summary).forEach(summaries::add);
    catalog.inspectionQueryTypes().stream().map(TypeEntry::summary).forEach(summaries::add);
    catalog.nestedTypes().stream()
        .map(NestedTypeGroup::types)
        .flatMap(List::stream)
        .map(TypeEntry::summary)
        .forEach(summaries::add);
    catalog.plainTypes().stream()
        .map(PlainTypeGroup::type)
        .map(TypeEntry::summary)
        .forEach(summaries::add);
    catalog.shippedExamples().stream().map(ShippedExampleEntry::summary).forEach(summaries::add);
    return String.join("\n", summaries);
  }

  private static Path repositoryRoot() {
    Path current = Path.of("").toAbsolutePath().normalize();
    while (current != null) {
      if (Files.isRegularFile(current.resolve("settings.gradle.kts"))) {
        return current;
      }
      current = current.getParent();
    }
    throw new IllegalStateException("Failed to locate repository root from " + Path.of(""));
  }
}
