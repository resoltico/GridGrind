package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises host acceptance directly across in-memory and materialized workbook execution. */
class HostAcceptanceExecutorTest {
  @TempDir Path root;

  @Test
  void verifiesTerminalAssertionsAndSemanticFactsForAnInMemoryWorkbook() throws Exception {
    HostAcceptanceExecutor executor = new HostAcceptanceExecutor(stepSupport());
    AcceptancePolicy.RequireAll policy =
        new AcceptancePolicy.RequireAll(
            List.of(
                terminalSheetAssertion(),
                preserveSheetSummary(),
                new AcceptancePolicy.Requirement.RequireCalculation()));

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      HostAcceptanceSnapshot snapshot =
          executor.captureBeforeMutation(
              policy, workbook, new WorkbookLocation.UnsavedWorkbook(), null);
      HostAcceptanceVerification verification =
          executor.verifyWorkbook(
              policy, snapshot, workbook, new WorkbookLocation.UnsavedWorkbook());

      assertEquals(1, verification.assertions().size());
      assertEquals(
          List.of("preserve-summary"),
          verification.preservation().stream()
              .map(WorkbookExecutionEvidence.Preservation.SemanticEstablished.class::cast)
              .flatMap(value -> value.inspectionStepIds().stream())
              .toList()
              .stream()
              .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
              .toList());
    }
  }

  @Test
  void verifiesTerminalAssertionsForAMaterializedWorkbook() throws Exception {
    Path workbookPath = writeWorkbook();
    HostAcceptanceExecutor executor = new HostAcceptanceExecutor(stepSupport());
    AcceptancePolicy.RequireAll policy =
        new AcceptancePolicy.RequireAll(List.of(terminalSheetAssertion()));
    HostAcceptanceVerification terminalVerification =
        executor.verifyMaterializedTerminalAcceptance(
            policy, workbookPath, new WorkbookLocation.UnsavedWorkbook());
    HostAcceptanceVerification minimumVerification =
        executor.verifyMaterializedTerminalAcceptance(
            AcceptancePolicy.minimum(), workbookPath, new WorkbookLocation.UnsavedWorkbook());
    HostAcceptanceVerification calculationOnlyVerification =
        executor.verifyMaterializedTerminalAcceptance(
            new AcceptancePolicy.RequireAll(
                List.of(new AcceptancePolicy.Requirement.RequireCalculation())),
            workbookPath,
            new WorkbookLocation.UnsavedWorkbook());

    assertEquals(1, terminalVerification.assertions().size());
    assertEquals(0, minimumVerification.assertions().size());
    assertEquals(0, calculationOnlyVerification.assertions().size());
  }

  @Test
  void refusesOpaquePreservationWithoutAMaterializedSourceAndKeepsMinimumPolicyEmpty()
      throws Exception {
    HostAcceptanceExecutor executor = new HostAcceptanceExecutor(stepSupport());
    AcceptancePolicy.RequireAll opaque =
        new AcceptancePolicy.RequireAll(
            List.of(
                new AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                    List.of("/customXml/item1.xml"))));

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      assertThrows(
          IOException.class,
          () ->
              executor.captureBeforeMutation(
                  opaque, workbook, new WorkbookLocation.UnsavedWorkbook(), null));
      assertEquals(
          List.of(),
          executor.verifyStagedArtifact(
              AcceptancePolicy.minimum(), HostAcceptanceSnapshot.empty(), root));
    }
  }

  @Test
  void establishesExplicitOpaquePartPreservationAgainstTheVerifiedStagedArtifact()
      throws Exception {
    HostAcceptanceExecutor executor = new HostAcceptanceExecutor(stepSupport());
    AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts requirement =
        new AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(List.of("/customXml/item1.xml"));
    AcceptancePolicy.RequireAll policy = new AcceptancePolicy.RequireAll(List.of(requirement));
    Path source = zipWithCustomPart("source.zip", "stable");
    Path staged = zipWithCustomPart("staged.zip", "stable");
    HostAcceptanceSnapshot snapshot =
        new HostAcceptanceSnapshot(
            List.of(),
            List.of(new HostAcceptanceSnapshot.OpaquePartPreservation(requirement, source)));

    assertEquals(
        List.of("/customXml/item1.xml"),
        assertInstanceOf(
                WorkbookExecutionEvidence.Preservation.ByteEstablished.class,
                executor.verifyStagedArtifact(policy, snapshot, staged).getFirst())
            .partNames());
  }

  @Test
  void snapshotRejectsUncapturedRequirementsAndMismatchedSemanticBaselines() {
    AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts opaque =
        new AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(List.of("/customXml/item1.xml"));
    AcceptancePolicy.Requirement.PreserveInspectionFacts semantic = preserveSheetSummary();

    assertThrows(
        IllegalStateException.class, () -> HostAcceptanceSnapshot.empty().sourceFor(opaque));
    assertThrows(
        IllegalStateException.class, () -> HostAcceptanceSnapshot.empty().baselineFor(semantic));
    assertThrows(
        IllegalArgumentException.class,
        () -> new HostAcceptanceSnapshot.SemanticPreservation(semantic, List.of()));
  }

  private AcceptancePolicy.Requirement.TerminalAssertions terminalSheetAssertion() {
    return new AcceptancePolicy.Requirement.TerminalAssertions(
        List.of(
            new AssertionStep(
                new StepId("assert-budget"),
                new SheetSelector.ByName(new SheetName("Budget")),
                new PresenceAssertion.SheetPresent())));
  }

  private AcceptancePolicy.Requirement.PreserveInspectionFacts preserveSheetSummary() {
    return new AcceptancePolicy.Requirement.PreserveInspectionFacts(
        List.of(
            new InspectionStep(
                new StepId("preserve-summary"),
                new SheetSelector.ByName(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetSheetSummary())));
  }

  private ExecutionStepSupport stepSupport() throws IOException {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    return new ExecutionStepSupport(
        workbookEngine,
        selectorResolver,
        new AssertionExecutor(workbookEngine, selectorResolver),
        WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile,
        new ExecutionInputBindings(
            root, root.resolve("scratch"), ExecutionGrantTestSupport.noPublication()));
  }

  private Path writeWorkbook() throws IOException {
    Path workbookPath = root.resolve("materialized.xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(workbookPath)) {
      workbook.createSheet("Budget");
      workbook.write(output);
    }
    return workbookPath;
  }

  private Path zipWithCustomPart(String fileName, String contents) throws IOException {
    Path archive = root.resolve(fileName);
    try (ZipOutputStream output = new ZipOutputStream(Files.newOutputStream(archive))) {
      output.putNextEntry(new ZipEntry("customXml/item1.xml"));
      output.write(contents.getBytes(java.nio.charset.StandardCharsets.UTF_8));
      output.closeEntry();
    }
    return archive;
  }
}
