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
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
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
    CalculationWorkflow calculationSupport = new CalculationWorkflow(writer -> {});
    ExecutionResults responseSupport = responseSupport();
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

    try (var workbook = XssfWorkbooks.create()) {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              workflow.execute(
                  new FullXssfWorkflow.Call(
                      ProtocolVersion.CURRENT,
                      MutationResolution.preparePlan(request),
                      workbook,
                      new ExecutionModeSpec.FullXssf(),
                      new ArrayList<>(),
                      ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                      bindings)));

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
    CalculationWorkflow calculationSupport = new CalculationWorkflow(writer -> {});
    ExecutionResults responseSupport = responseSupport();
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

    try (var workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("A1+1"));

      assertInstanceOf(
          WorkbookResult.Failure.class,
          workflow.execute(
              new FullXssfWorkflow.Call(
                  ProtocolVersion.CURRENT,
                  MutationResolution.preparePlan(request),
                  workbook,
                  new ExecutionModeSpec.FullXssf(),
                  new ArrayList<>(),
                  ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                  bindings)));
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
            new HostAcceptanceExecutor(stepSupport()), new CalculationWorkflow(writer -> {}));

    try (var workbook = XssfWorkbooks.create()) {
      FullXssfWorkflowState state =
          new FullXssfWorkflowState(
              ProtocolVersion.CURRENT,
              MutationResolution.preparePlan(request),
              workbook,
              new ArrayList<>(),
              ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
              bindings);

      acceptance.recordFailure(state, new java.io.IOException("storage unavailable"));

      assertEquals(List.of(), state.hostAssertions());
      assertEquals(List.of(), state.hostPreservation());
    }
  }

  private StepExecution stepSupport() {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    return new StepExecution(
        workbookEngine,
        selectorResolver,
        new AssertionExecutor(workbookEngine, selectorResolver),
        WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile,
        new ExecutionInputBindings(
            root, root.resolve("scratch"), ExecutionGrantTestSupport.noPublication()));
  }

  private static ExecutionResults responseSupport() {
    return new ExecutionResults(
        dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook::close,
        materialized -> materialized.close());
  }
}
