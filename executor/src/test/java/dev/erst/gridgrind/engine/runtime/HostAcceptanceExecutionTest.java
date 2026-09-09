package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellRowInput;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Proves trusted-host acceptance executes after mutations and before final workbook publication.
 */
class HostAcceptanceExecutionTest {
  @TempDir Path root;

  @Test
  void failingHostAcceptanceBlocksPublicationEvenWhenTheAuthoredPlanHasNoAssertion() {
    Path output = root.resolve("approved-report.xlsx");
    WorkbookPlan request = reportPlan(output);
    GridGrindHostAcceptancePolicy policy =
        new GridGrindHostAcceptancePolicy.RequireAll(
            List.of(
                new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(
                    List.of(
                        new AssertionStep(
                            "host-total",
                            new CellSelector.ByAddress("Report", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.NumberValue(2)))))));

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(root, request, policy),
                    ExecutionProgressSink.NOOP));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.TaskSpecific.Failed evidence =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.TaskSpecific.Failed.class,
            failure.evidence().taskSpecific());
    assertFalse(Files.exists(output));
    assertEquals(
        List.of(
            dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.TaskSpecific.Origin
                .HOST_REQUIRED),
        evidence.assertions().stream().map(value -> value.origin()).toList());
  }

  @Test
  void successfulHostAcceptanceIsReportedAsHostRequiredEvidence() {
    Path output = root.resolve("approved-report.xlsx");
    WorkbookPlan request = reportPlan(output);
    GridGrindHostAcceptancePolicy policy =
        new GridGrindHostAcceptancePolicy.RequireAll(
            List.of(
                new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(
                    List.of(
                        new AssertionStep(
                            "host-total",
                            new CellSelector.ByAddress("Report", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.NumberValue(1)))))));

    WorkbookResult.Success success =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(root, request, policy),
                    ExecutionProgressSink.NOOP));

    dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.TaskSpecific.Established evidence =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.TaskSpecific.Established
                .class,
            success.evidence().taskSpecific());
    assertTrue(Files.isRegularFile(output));
    assertEquals(
        List.of(
            dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.TaskSpecific.Origin
                .HOST_REQUIRED),
        evidence.assertions().stream().map(value -> value.origin()).toList());
  }

  @Test
  void hostRequiredCalculationRunsBeforeFullXssfPublication() {
    Path output = root.resolve("calculated-report.xlsx");
    WorkbookPlan request = reportPlan(output);
    GridGrindHostAcceptancePolicy policy =
        new GridGrindHostAcceptancePolicy.RequireAll(
            List.of(new GridGrindHostAcceptancePolicy.Requirement.RequireCalculation()));

    WorkbookResult.Success success =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(root, request, policy),
                    ExecutionProgressSink.NOOP));

    assertTrue(Files.isRegularFile(output));
    assertNotEquals(
        dev.erst.gridgrind.contract.dto.CalculationExecutionStatus.NOT_REQUESTED,
        success.calculation().execution().status());
  }

  @Test
  void streamingWriteEvaluatesHostAcceptanceBeforePublication() {
    Path output = root.resolve("streaming-report.xlsx");
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                output.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicyInput.mode(new ExecutionModeInput.StreamingWrite()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "ensure-report",
                    new SheetSelector.ByName("Report"),
                    new WorkbookMutationAction.EnsureSheet()),
                new MutationStep(
                    "append-total",
                    new SheetSelector.ByName("Report"),
                    new CellMutationAction.AppendRow(
                        new CellRowInput.Typed(List.of(new CellInput.NumberValue(1)))))));

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(
                        root, request, failingHostTotal()),
                    ExecutionProgressSink.NOOP));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertFalse(Files.exists(output));
  }

  @Test
  void directEventReadEvaluatesHostAcceptanceWithoutBypassingAdmission() throws Exception {
    Path source = root.resolve("source.xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet("Report").createRow(0).createCell(0).setCellValue(1);
      workbook.write(output);
    }
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(new ExecutionModeInput.EventRead()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new InspectionStep(
                    "summary",
                    new SheetSelector.ByName("Report"),
                    new SheetIntrospectionQuery.GetSheetSummary())));

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(
                        root, request, failingHostTotal()),
                    ExecutionProgressSink.NOOP));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
  }

  @Test
  void hostSemanticPreservationIsEstablishedFromBeforeAndAfterFacts() throws Exception {
    Path source = writePreservationSource();
    Path output = root.resolve("preserved.xlsx");
    WorkbookPlan request = preservationPlan(source, output, "A1", 2);

    WorkbookResult result =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request,
                ExecutionInputBindingsFixtureSupport.bindings(root, request, preserveCell("B1")),
                ExecutionProgressSink.NOOP);
    WorkbookResult.Success success =
        assertInstanceOf(WorkbookResult.Success.class, result, () -> result.toString());

    dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation.SemanticEstablished
        preservation =
            assertInstanceOf(
                dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation
                    .SemanticEstablished.class,
                success.evidence().preservation().getFirst());
    assertEquals(List.of("preserve-unaffected-cell"), preservation.inspectionStepIds());
  }

  @Test
  void failedHostSemanticPreservationBlocksPublication() throws Exception {
    Path source = writePreservationSource();
    Path output = root.resolve("preserved.xlsx");
    WorkbookPlan request = preservationPlan(source, output, "B1", 2);

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(
                        root, request, preserveCell("B1")),
                    ExecutionProgressSink.NOOP));

    assertEquals(GridGrindProblemCode.PRESERVATION_FAILED, failure.problem().code());
    assertFalse(Files.exists(output));
    dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation.Failed preservation =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation.Failed.class,
            failure.evidence().preservation().getFirst());
    assertEquals(
        dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation.Kind.SEMANTIC,
        preservation.kind());
  }

  private static WorkbookPlan reportPlan(Path output) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            output.toString(),
            WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
            dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none()),
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
        reportSteps());
  }

  private static List<WorkbookStep> reportSteps() {
    return List.of(
        new MutationStep(
            "ensure-report",
            new SheetSelector.ByName("Report"),
            new WorkbookMutationAction.EnsureSheet()),
        new MutationStep(
            "write-total",
            new CellSelector.ByAddress("Report", "A1"),
            new CellMutationAction.SetCell(new CellInput.NumberValue(1))));
  }

  private static GridGrindHostAcceptancePolicy failingHostTotal() {
    return new GridGrindHostAcceptancePolicy.RequireAll(
        List.of(
            new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(
                List.of(
                    new AssertionStep(
                        "host-total",
                        new CellSelector.ByAddress("Report", "A1"),
                        new CellAssertion.CellValue(new CellScalarValue.NumberValue(2)))))));
  }

  private Path writePreservationSource() throws java.io.IOException {
    Path source = root.resolve("preservation-source.xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      var row = workbook.createSheet("Report").createRow(0);
      row.createCell(0).setCellValue(1);
      row.createCell(1).setCellValue(1);
      workbook.write(output);
    }
    return source;
  }

  private static WorkbookPlan preservationPlan(
      Path source, Path output, String address, double value) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            output.toString(),
            WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
            dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none()),
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
        List.of(
            new MutationStep(
                "update-report",
                new CellSelector.ByAddress("Report", address),
                new CellMutationAction.SetCell(new CellInput.NumberValue(value)))));
  }

  private static GridGrindHostAcceptancePolicy preserveCell(String address) {
    return new GridGrindHostAcceptancePolicy.RequireAll(
        List.of(
            new GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts(
                List.of(
                    new InspectionStep(
                        "preserve-unaffected-cell",
                        new CellSelector.ByAddress("Report", address),
                        new SheetIntrospectionQuery.GetCells())))));
  }
}
