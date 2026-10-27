package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.PreparedMutationStep;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.ResolvedTarget;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ResolvedMutationExecutor;
import dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter;
import java.io.IOException;
import java.util.Objects;

/**
 * Resolves one mutation, validates normal formulas in workbook state, then records their origin.
 */
final class FormulaAwareMutationStepExecutor {
  private final SemanticSelectorResolver selectorResolver;
  private final ExecutionInputBindings bindings;

  FormulaAwareMutationStepExecutor(
      SemanticSelectorResolver selectorResolver, ExecutionInputBindings bindings) {
    this.selectorResolver =
        Objects.requireNonNull(selectorResolver, "selectorResolver must not be null");
    this.bindings = Objects.requireNonNull(bindings, "bindings must not be null");
  }

  void execute(
      ExcelWorkbook workbook,
      PreparedMutationStep prepared,
      FormulaOriginTracker formulaOrigins,
      StepReference authoringStep)
      throws IOException {
    MutationStep mutationStep = prepared.authoredStep();
    ResolvedStep step =
        new ResolvedStep(
            (ResolvedTarget)
                selectorResolver.resolveMutationTarget(
                    workbook, mutationStep.target(), mutationStep.action()),
            MutationResolution.finish(prepared.action(), bindings));
    FormulaOriginTracker.FormulaWrites writes = formulaOrigins.plannedWrites(workbook, step);
    ResolvedMutationExecutor.apply(workbook, step);
    formulaOrigins.record(workbook, writes, authoringStep);
  }

  void executeStreaming(ExcelStreamingWorkbookWriter writer, PreparedMutationStep prepared)
      throws IOException {
    MutationStep mutationStep = prepared.authoredStep();
    writer.apply(
        new ResolvedStep(
            (ResolvedTarget) mutationStep.target(),
            MutationResolution.finish(prepared.action(), bindings)));
  }
}
