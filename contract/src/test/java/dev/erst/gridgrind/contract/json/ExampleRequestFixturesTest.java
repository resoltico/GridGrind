package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Parses shipped example requests so public fixtures stay aligned with the step schema. */
final class ExampleRequestFixturesTest {
  @Test
  void repositoryExamplesExactlyMatchTheGeneratedContractRegistry() throws IOException {
    List<Path> exampleFiles;
    try (Stream<Path> stream = Files.list(examplesDirectory())) {
      exampleFiles =
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .sorted(Comparator.comparing(path -> path.getFileName().toString()))
              .toList();
    }

    List<String> actualFiles =
        exampleFiles.stream().map(path -> path.getFileName().toString()).toList();
    List<String> expectedFiles =
        GridGrindShippedExamples.repositoryExamples().stream()
            .map(GridGrindShippedExamples.ShippedExample::fileName)
            .sorted()
            .toList();

    assertEquals(
        expectedFiles, actualFiles, "examples/ must mirror the generated registry exactly");
    for (GridGrindShippedExamples.ShippedExample example :
        GridGrindShippedExamples.repositoryExamples()) {
      Path exampleFile = examplesDirectory().resolve(example.fileName());
      byte[] expectedBytes = GridGrindJson.writeRequestBytes(example.plan());
      byte[] actualBytes = Files.readAllBytes(exampleFile);
      WorkbookPlan request =
          assertDoesNotThrow(
              () -> GridGrindJson.readRequest(actualBytes),
              () -> "Failed to parse request example " + exampleFile);
      assertNotNull(request, () -> "Parsed request example must not be null: " + exampleFile);
      assertEquals(
          example.plan(),
          request,
          () -> "Parsed example must match the generated plan: " + exampleFile);
      assertEquals(
          expectedFixtureText(expectedBytes),
          fixtureText(actualBytes),
          () -> "Example fixture drifted from generated plan bytes: " + exampleFile);
      assertFalse(
          fixtureText(actualBytes).contains(": null"),
          () -> "Example fixture must omit explicit null properties: " + exampleFile);
    }
  }

  @Test
  void workbookHealthExampleDemonstratesNoSaveBatchAnalysisWorkflow() throws IOException {
    WorkbookPlan request = readExample("workbook-health-request.json");

    assertInstanceOf(WorkbookPlan.WorkbookPersistence.None.class, request.persistence());
    assertTrue(request.steps().stream().anyMatch(MutationStep.class::isInstance));
    assertTrue(
        request.steps().stream()
            .filter(InspectionStep.class::isInstance)
            .map(InspectionStep.class::cast)
            .map(InspectionStep::query)
            .anyMatch(InspectionQuery.GetSheetSummary.class::isInstance));
    assertTrue(
        request.steps().stream()
            .filter(InspectionStep.class::isInstance)
            .map(InspectionStep.class::cast)
            .map(InspectionStep::query)
            .anyMatch(InspectionQuery.AnalyzeFormulaHealth.class::isInstance));
    assertTrue(
        request.steps().stream()
            .filter(InspectionStep.class::isInstance)
            .map(InspectionStep.class::cast)
            .map(InspectionStep::query)
            .anyMatch(InspectionQuery.AnalyzeWorkbookFindings.class::isInstance));
  }

  @Test
  void assertionExampleDemonstratesMutateThenVerifyWorkflow() throws IOException {
    WorkbookPlan request = readExample("assertion-request.json");

    assertEquals(ExecutionJournalLevel.VERBOSE, request.journalLevel());
    assertTrue(request.steps().stream().anyMatch(MutationStep.class::isInstance));
    assertTrue(request.steps().stream().anyMatch(AssertionStep.class::isInstance));
    assertTrue(
        request.steps().stream()
            .filter(AssertionStep.class::isInstance)
            .map(AssertionStep.class::cast)
            .map(AssertionStep::assertion)
            .anyMatch(
                assertion -> "EXPECT_ANALYSIS_MAX_SEVERITY".equals(assertion.assertionType())));
    assertTrue(
        request.steps().stream()
            .filter(InspectionStep.class::isInstance)
            .map(InspectionStep.class::cast)
            .map(InspectionStep::query)
            .anyMatch(InspectionQuery.GetCells.class::isInstance));
  }

  private static WorkbookPlan readExample(String fileName) throws IOException {
    return GridGrindJson.readRequest(Files.readAllBytes(examplesDirectory().resolve(fileName)));
  }

  private static String expectedFixtureText(byte[] requestBytes) {
    return fixtureText(requestBytes) + System.lineSeparator();
  }

  private static String fixtureText(byte[] bytes) {
    return new String(bytes, StandardCharsets.UTF_8);
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
