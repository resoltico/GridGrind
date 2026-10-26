package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
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
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
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
            new ExecutionGrant(
                List.of(),
                List.of(),
                new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                new AcceptancePolicy.RequireAll(
                    List.of(
                        new AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                            List.of("/customXml/item1.xml"))))));

    try (var workbook = ExcelWorkbooks.create()) {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              workflow.execute(
                  ProtocolVersion.CURRENT,
                  request,
                  workbook,
                  new ExecutionModeSpec.FullXssf(),
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
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    AcceptancePolicy policy =
        new AcceptancePolicy.RequireAll(
            List.of(new AcceptancePolicy.Requirement.RequireCalculation()));
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
            new ExecutionGrant(
                List.of(),
                List.of(),
                new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                policy));

    try (var workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("A1+1"));

      assertInstanceOf(
          WorkbookResult.Failure.class,
          workflow.execute(
              ProtocolVersion.CURRENT,
              request,
              workbook,
              new ExecutionModeSpec.FullXssf(),
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
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
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
              ProtocolVersion.CURRENT,
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
