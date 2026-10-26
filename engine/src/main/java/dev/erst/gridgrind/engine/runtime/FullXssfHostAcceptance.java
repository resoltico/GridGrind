package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Applies the host-owned conditions that a full-XSSF plan cannot weaken or bypass. */
final class FullXssfHostAcceptance {
  private final HostAcceptanceExecutor acceptanceExecutor;
  private final ExecutionCalculationSupport calculationSupport;

  FullXssfHostAcceptance(
      HostAcceptanceExecutor acceptanceExecutor, ExecutionCalculationSupport calculationSupport) {
    this.acceptanceExecutor = acceptanceExecutor;
    this.calculationSupport = calculationSupport;
  }

  Context capture(FullXssfWorkflowState state, ExecutionInputBindings bindings) throws IOException {
    AcceptancePolicy policy = bindings.executionGrant().acceptancePolicy();
    state.hostAcceptanceSnapshot(
        acceptanceExecutor.captureBeforeMutation(
            policy,
            state.executionContext().workbook(),
            state.workbookLocation(),
            materializedSourceArtifact(state.executionContext().request(), bindings)));
    return new Context(policy);
  }

  Optional<dev.erst.gridgrind.core.problem.GridGrindProblemDetail.Problem> requireCalculation(
      FullXssfWorkflowState state, Context context) {
    if (!requiresStrictCalculation(context.policy())) {
      return Optional.empty();
    }
    var policy =
        new CalculationPolicySpec(
            new CalculationStrategySpec.RequireEvaluation(),
            state.executionContext().request().calculationPolicy().markRecalculateOnOpen());
    ExecutionCalculationSupport.CalculationExecutionOutcome outcome =
        calculationSupport.executeCalculationPolicy(
            state.executionContext().workbook(),
            state.executionContext().request(),
            policy,
            state.executionContext().journal(),
            state.formulaOrigins(),
            null);
    state.calculation(outcome.report());
    state.executionContext().warnings().addAll(outcome.warnings());
    return outcome.failure();
  }

  void verify(FullXssfWorkflowState state, Context context)
      throws IOException, AssertionFailedException, PreservationFailedException {
    HostAcceptanceVerification verification =
        acceptanceExecutor.verifyWorkbook(
            context.policy(),
            state.hostAcceptanceSnapshot(),
            state.executionContext().workbook(),
            state.workbookLocation());
    state.hostAssertions().addAll(verification.assertions());
    state.hostPreservation().addAll(verification.preservation());
  }

  StagedArtifactAcceptance stagedArtifactAcceptance(FullXssfWorkflowState state, Context context) {
    return stagedWorkbook ->
        state
            .hostPreservation()
            .addAll(
                acceptanceExecutor.verifyStagedArtifact(
                    context.policy(), state.hostAcceptanceSnapshot(), stagedWorkbook));
  }

  void recordFailure(FullXssfWorkflowState state, Exception exception) {
    switch (exception) {
      case PreservationFailedException preservationFailed ->
          state
              .hostPreservation()
              .add(
                  new Preservation.Failed(
                      preservationFailed.kind(), preservationFailed.identifiers()));
      case AssertionFailedException assertionFailed ->
          state
              .hostAssertions()
              .add(
                  new AssertionResult.Failed(
                      assertionFailed.assertionFailure().stepId(),
                      assertionFailed.assertionFailure().assertionType(),
                      assertionFailed.assertionFailure()));
      default -> {}
    }
  }

  private static @Nullable Path materializedSourceArtifact(
      WorkbookPlan request, ExecutionInputBindings bindings) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> null;
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          bindings.requestPathAccess().materializedReadPath(existingFile.path());
    };
  }

  private static boolean requiresStrictCalculation(AcceptancePolicy policy) {
    return switch (policy) {
      case AcceptancePolicy.MinimumOnly _ -> false;
      case AcceptancePolicy.RequireAll requireAll ->
          requireAll.requirements().stream()
              .anyMatch(AcceptancePolicy.Requirement.RequireCalculation.class::isInstance);
    };
  }

  record Context(AcceptancePolicy policy) {}
}
