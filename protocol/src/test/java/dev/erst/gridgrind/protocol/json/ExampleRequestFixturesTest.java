package dev.erst.gridgrind.protocol.json;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Parses shipped example requests so public fixtures stay aligned with the request schema. */
final class ExampleRequestFixturesTest {
  @Test
  void allShippedExampleRequestsParse() throws IOException {
    List<Path> exampleFiles;
    try (Stream<Path> stream = Files.list(examplesDirectory())) {
      exampleFiles =
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .sorted(Comparator.comparing(path -> path.getFileName().toString()))
              .toList();
    }

    assertFalse(exampleFiles.isEmpty(), "Expected shipped request examples under examples/");
    for (Path exampleFile : exampleFiles) {
      GridGrindRequest request =
          assertDoesNotThrow(
              () -> GridGrindJson.readRequest(Files.readAllBytes(exampleFile)),
              () -> "Failed to parse request example " + exampleFile);
      assertNotNull(request, () -> "Parsed request example must not be null: " + exampleFile);
    }
  }

  @Test
  void introspectionAnalysisExampleDemonstratesDistinctReadAndAnalysisOperations()
      throws IOException {
    GridGrindRequest request = readExample("introspection-analysis-request.json");

    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.GetFormulaSurface.class::isInstance));
    assertTrue(
        request.reads().stream().anyMatch(WorkbookReadOperation.GetSheetSchema.class::isInstance));
    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.GetNamedRangeSurface.class::isInstance));
    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.AnalyzeFormulaHealth.class::isInstance));
    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.AnalyzeHyperlinkHealth.class::isInstance));
    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.AnalyzeNamedRangeHealth.class::isInstance));
    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.AnalyzeWorkbookFindings.class::isInstance));
  }

  @Test
  void workbookHealthExampleDemonstratesNoSaveBatchAnalysisWorkflow() throws IOException {
    GridGrindRequest request = readExample("workbook-health-request.json");

    assertInstanceOf(GridGrindRequest.WorkbookPersistence.None.class, request.persistence());
    assertTrue(
        request.reads().stream().anyMatch(WorkbookReadOperation.GetSheetSummary.class::isInstance));
    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.AnalyzeFormulaHealth.class::isInstance));
    assertTrue(
        request.reads().stream()
            .anyMatch(WorkbookReadOperation.AnalyzeWorkbookFindings.class::isInstance));
  }

  private static GridGrindRequest readExample(String fileName) throws IOException {
    return GridGrindJson.readRequest(Files.readAllBytes(examplesDirectory().resolve(fileName)));
  }

  private static Path examplesDirectory() {
    Path candidate = Path.of("").toAbsolutePath().normalize();
    while (candidate != null) {
      if (Files.exists(candidate.resolve("gradle.properties"))
          && Files.exists(candidate.resolve("examples"))) {
        return candidate.resolve("examples");
      }
      candidate = candidate.getParent();
    }
    throw new AssertionError("Could not locate the GridGrind examples directory.");
  }
}
