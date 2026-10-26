package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
                            new StepId("host-total"),
                            new CellSelector.ByAddress(
                                new SheetName("Report"), new CellAddress("A1")),
                            new CellAssertion.Value(new CellValue.NumberValue(2)))))));

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(root, request, policy),
                    ExecutionProgressSink.NOOP));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Failed evidence =
        assertInstanceOf(
            dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Failed.class,
            failure.evidence().taskSpecific());
    assertFalse(Files.exists(output));
    assertEquals(
        List.of(
            dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Origin
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
                            new StepId("host-total"),
                            new CellSelector.ByAddress(
                                new SheetName("Report"), new CellAddress("A1")),
                            new CellAssertion.Value(new CellValue.NumberValue(1)))))));

    WorkbookResult.Success success =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request,
                    ExecutionInputBindingsFixtureSupport.bindings(root, request, policy),
                    ExecutionProgressSink.NOOP));

    dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Established evidence =
        assertInstanceOf(
            dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Established.class,
            success.evidence().taskSpecific());
    assertTrue(Files.isRegularFile(output));
    assertEquals(
        List.of(
            dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.TaskSpecific.Origin
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
        dev.erst.gridgrind.core.model.formula.CalculationExecutionStatus.NOT_REQUESTED,
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
            ExecutionPolicySpec.mode(new ExecutionModeSpec.StreamingWrite()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure-report"),
                    new SheetSelector.ByName(new SheetName("Report")),
                    new WorkbookMutationAction.EnsureSheet()),
                new MutationStep(
                    new StepId("append-total"),
                    new SheetSelector.ByName(new SheetName("Report")),
                    new CellMutationAction.AppendRow(
                        new CellRowSpec.Typed(List.of(new CellValueSpec.NumberValue(1)))))));

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
            ExecutionPolicySpec.mode(new ExecutionModeSpec.EventRead()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new InspectionStep(
                    new StepId("summary"),
                    new SheetSelector.ByName(new SheetName("Report")),
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

    dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation.SemanticEstablished
        preservation =
            assertInstanceOf(
                dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation
                    .SemanticEstablished.class,
                success.evidence().preservation().getFirst());
    assertEquals(
        List.of("preserve-unaffected-cell"),
        preservation.inspectionStepIds().stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
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
    dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation.Failed preservation =
        assertInstanceOf(
            dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation.Failed.class,
            failure.evidence().preservation().getFirst());
    assertEquals(
        dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation.Kind.SEMANTIC,
        preservation.kind());
  }

  private static WorkbookPlan reportPlan(Path output) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            output.toString(),
            WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
            dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        reportSteps());
  }

  private static List<WorkbookStep> reportSteps() {
    return List.of(
        new MutationStep(
            new StepId("ensure-report"),
            new SheetSelector.ByName(new SheetName("Report")),
            new WorkbookMutationAction.EnsureSheet()),
        new MutationStep(
            new StepId("write-total"),
            new CellSelector.ByAddress(new SheetName("Report"), new CellAddress("A1")),
            new CellMutationAction.SetCell(new CellValueSpec.NumberValue(1))));
  }

  private static GridGrindHostAcceptancePolicy failingHostTotal() {
    return new GridGrindHostAcceptancePolicy.RequireAll(
        List.of(
            new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(
                List.of(
                    new AssertionStep(
                        new StepId("host-total"),
                        new CellSelector.ByAddress(new SheetName("Report"), new CellAddress("A1")),
                        new CellAssertion.Value(new CellValue.NumberValue(2)))))));
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
            dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new MutationStep(
                new StepId("update-report"),
                new CellSelector.ByAddress(new SheetName("Report"), new CellAddress(address)),
                new CellMutationAction.SetCell(new CellValueSpec.NumberValue(value)))));
  }

  private static GridGrindHostAcceptancePolicy preserveCell(String address) {
    return new GridGrindHostAcceptancePolicy.RequireAll(
        List.of(
            new GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts(
                List.of(
                    new InspectionStep(
                        new StepId("preserve-unaffected-cell"),
                        new CellSelector.ByAddress(
                            new SheetName("Report"), new CellAddress(address)),
                        new SheetIntrospectionQuery.GetCells())))));
  }
}
