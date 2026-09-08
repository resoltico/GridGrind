package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies full-XSSF closes safely when host acceptance cannot capture its required baseline. */
class FullXssfWorkflowTest {
  @TempDir Path root;

  @Test
  void returnsStructuredFailureWhenOpaqueHostAcceptanceHasNoMaterializedSourceArtifact()
      throws Exception {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    ExecutionCalculationSupport calculationSupport = new ExecutionCalculationSupport(writer -> {});
    ExecutionResponseSupport responseSupport = responseSupport();
    FullXssfWorkflow workflow =
        new FullXssfWorkflow(
            calculationSupport,
            new FullXssfStepLoop(null, null),
            new FullXssfPersistence(
                (workbook, source, persistence, bindings, acceptance) -> {
                  throw new AssertionError("host capture must fail before persistence");
                },
                responseSupport),
            new FullXssfHostAcceptance(
                new HostAcceptanceExecutor(stepSupport()), calculationSupport),
            responseSupport);
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            new GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of(),
                new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new GridGrindExecutionGrant.PublicationAuthority.None(),
                List.of(),
                new GridGrindHostAcceptancePolicy.RequireAll(
                    List.of(
                        new GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                            List.of("/customXml/item1.xml"))))));

    try (var workbook = ExcelWorkbooks.create()) {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              workflow.execute(
                  GridGrindProtocolVersion.current(),
                  request,
                  workbook,
                  new ExecutionModeInput.FullXssf(),
                  new ArrayList<>(),
                  ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                  bindings));

      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    }
  }

  @Test
  void closesBeforePublicationWhenMandatoryHostCalculationCannotBeEstablished() throws Exception {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    GridGrindHostAcceptancePolicy policy =
        new GridGrindHostAcceptancePolicy.RequireAll(
            List.of(new GridGrindHostAcceptancePolicy.Requirement.RequireCalculation()));
    ExecutionCalculationSupport calculationSupport = new ExecutionCalculationSupport(writer -> {});
    ExecutionResponseSupport responseSupport = responseSupport();
    FullXssfWorkflow workflow =
        new FullXssfWorkflow(
            calculationSupport,
            new FullXssfStepLoop(null, null),
            new FullXssfPersistence(
                (workbook, source, persistence, bindings, acceptance) -> {
                  throw new AssertionError(
                      "mandatory calculation failure must prevent persistence");
                },
                responseSupport),
            new FullXssfHostAcceptance(
                new HostAcceptanceExecutor(stepSupport()), calculationSupport),
            responseSupport);
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            new GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of(),
                new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new GridGrindExecutionGrant.PublicationAuthority.None(),
                List.of(),
                policy));

    try (var workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", ExcelCellValue.formula("A1+1"));

      assertInstanceOf(
          WorkbookResult.Failure.class,
          workflow.execute(
              GridGrindProtocolVersion.current(),
              request,
              workbook,
              new ExecutionModeInput.FullXssf(),
              new ArrayList<>(),
              ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
              bindings));
    }
  }

  @Test
  void doesNotMisrepresentGenericIoFailureAsAnAcceptanceFinding() throws Exception {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root, root.resolve("scratch"), ExecutionGrantTestSupport.noPublication());
    FullXssfHostAcceptance acceptance =
        new FullXssfHostAcceptance(
            new HostAcceptanceExecutor(stepSupport()),
            new ExecutionCalculationSupport(writer -> {}));

    try (var workbook = ExcelWorkbooks.create()) {
      FullXssfWorkflowState state =
          new FullXssfWorkflowState(
              GridGrindProtocolVersion.current(),
              request,
              workbook,
              new ArrayList<>(),
              ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
              bindings);

      acceptance.recordFailure(state, new java.io.IOException("storage unavailable"));

      assertEquals(List.of(), state.hostAssertions());
      assertEquals(List.of(), state.hostPreservation());
    }
  }

  private ExecutionStepSupport stepSupport() {
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

  private static ExecutionResponseSupport responseSupport() {
    return new ExecutionResponseSupport(
        dev.erst.gridgrind.excel.ExcelWorkbook::close, materialized -> materialized.close());
  }
}
