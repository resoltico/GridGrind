package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
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
    GridGrindHostAcceptancePolicy policy =
        ((GridGrindExecutionGrant.Bounded) bindings.executionGrant()).hostAcceptancePolicy();
    state.hostAcceptanceSnapshot(
        acceptanceExecutor.captureBeforeMutation(
            policy,
            state.executionContext().workbook(),
            state.workbookLocation(),
            materializedSourceArtifact(state.executionContext().request(), bindings)));
    return new Context(policy);
  }

  Optional<dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem> requireCalculation(
      FullXssfWorkflowState state, Context context) {
    if (!requiresStrictCalculation(context.policy())) {
      return Optional.empty();
    }
    var policy =
        new CalculationPolicyInput(
            new CalculationStrategyInput.RequireEvaluation(),
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

  private static boolean requiresStrictCalculation(GridGrindHostAcceptancePolicy policy) {
    return switch (policy) {
      case GridGrindHostAcceptancePolicy.MinimumOnly _ -> false;
      case GridGrindHostAcceptancePolicy.RequireAll requireAll ->
          requireAll.requirements().stream()
              .anyMatch(
                  GridGrindHostAcceptancePolicy.Requirement.RequireCalculation.class::isInstance);
    };
  }

  record Context(GridGrindHostAcceptancePolicy policy) {}
}
