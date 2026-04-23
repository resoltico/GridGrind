package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import org.junit.jupiter.api.Test;

/** Build-failing coverage audit over the split public reference docs. */
class GridGrindReferenceDocsCoverageTest {
  private static final Pattern THIRD_LEVEL_ID_HEADING = Pattern.compile("(?m)^### ([A-Z0-9_]+)$");
  private static final Pattern ASSERTION_TABLE_ROW = Pattern.compile("(?m)^\\| `([A-Z0-9_]+)` \\|");

  @Test
  void mutationReferenceHeadingsMatchCatalogMutationIds() throws IOException {
    Set<String> expected =
        GridGrindProtocolCatalog.catalog().mutationActionTypes().stream()
            .map(TypeEntry::id)
            .collect(Collectors.toCollection(LinkedHashSet::new));
    List<String> actualHeadings = new ArrayList<>();
    for (String relativePath :
        List.of(
            "docs/WORKBOOK_AND_SHEET_MUTATIONS.md",
            "docs/LAYOUT_AND_STRUCTURE_MUTATIONS.md",
            "docs/CELL_VALUE_MUTATIONS.md",
            "docs/LINK_AND_COMMENT_MUTATIONS.md",
            "docs/DRAWING_MUTATIONS.md",
            "docs/STYLE_AND_VALIDATION_MUTATIONS.md",
            "docs/STRUCTURED_DATA_MUTATIONS.md")) {
      actualHeadings.addAll(thirdLevelIdHeadings(readDoc(relativePath)));
    }

    assertEquals(
        actualHeadings.size(),
        new LinkedHashSet<>(actualHeadings).size(),
        "Mutation reference docs must not duplicate mutation headings");
    assertEquals(
        expected,
        new LinkedHashSet<>(actualHeadings),
        "Mutation reference docs must cover every catalog mutation action exactly once");
  }

  @Test
  void inspectionReferenceHeadingsMatchCatalogInspectionQueryIds() throws IOException {
    Set<String> expected =
        GridGrindProtocolCatalog.catalog().inspectionQueryTypes().stream()
            .map(TypeEntry::id)
            .collect(Collectors.toCollection(LinkedHashSet::new));
    List<String> actualHeadings = new ArrayList<>();
    for (String relativePath :
        List.of(
            "docs/WORKBOOK_AND_CELL_INSPECTIONS.md",
            "docs/DRAWING_AND_STRUCTURED_INSPECTIONS.md",
            "docs/ANALYSIS_QUERIES.md")) {
      actualHeadings.addAll(thirdLevelIdHeadings(readDoc(relativePath)));
    }

    assertEquals(
        actualHeadings.size(),
        new LinkedHashSet<>(actualHeadings).size(),
        "Assertion and inspection reference must not duplicate inspection headings");
    assertEquals(
        expected,
        new LinkedHashSet<>(actualHeadings),
        "Assertion and inspection reference must cover every catalog inspection query exactly once");
  }

  @Test
  void assertionFamiliesTableMatchesCatalogAssertionIds() throws IOException {
    Set<String> expected =
        GridGrindProtocolCatalog.catalog().assertionTypes().stream()
            .map(TypeEntry::id)
            .collect(Collectors.toCollection(LinkedHashSet::new));
    String document = readDoc("docs/ASSERTIONS.md");
    String assertionTable =
        between(document, "Assertion families:", "Common assertion step shapes:");
    List<String> actualRows = findAll(ASSERTION_TABLE_ROW, assertionTable);

    assertEquals(
        actualRows.size(),
        new LinkedHashSet<>(actualRows).size(),
        "Assertion families table must not duplicate assertion rows");
    assertEquals(
        expected,
        new LinkedHashSet<>(actualRows),
        "Assertion families table must cover every catalog assertion type exactly once");
  }

  @Test
  void publicEntryDocsLinkToTheJavaAuthoringGuide() throws IOException {
    assertTrue(
        readDoc("README.md").contains("docs/JAVA_AUTHORING.md"),
        "README must link to the Java authoring guide");
    assertTrue(
        readDoc("docs/QUICK_START.md").contains("JAVA_AUTHORING.md"),
        "QUICK_START must link to the Java authoring guide");
    assertTrue(
        readDoc("docs/OPERATIONS.md").contains("JAVA_AUTHORING.md"),
        "OPERATIONS must link to the Java authoring guide");
    assertTrue(
        readDoc("docs/JAVA_AUTHORING.md").contains("examples/java-authoring-workflow.java"),
        "Java authoring guide must point at the compile-verified example");
  }

  private static List<String> thirdLevelIdHeadings(String document) {
    return findAll(THIRD_LEVEL_ID_HEADING, document);
  }

  private static List<String> findAll(Pattern pattern, String text) {
    List<String> matches = new ArrayList<>();
    Matcher matcher = pattern.matcher(text);
    while (matcher.find()) {
      matches.add(matcher.group(1));
    }
    return List.copyOf(matches);
  }

  private static String between(String document, String start, String end) {
    int startIndex = document.indexOf(start);
    int endIndex = document.indexOf(end);
    if (startIndex < 0 || endIndex < 0 || endIndex <= startIndex) {
      throw new AssertionError(
          "Could not isolate doc section between '" + start + "' and '" + end + "'");
    }
    return document.substring(startIndex, endIndex);
  }

  private static String readDoc(String relativePath) throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    return Files.readString(repositoryRoot.resolve(relativePath));
  }
}
