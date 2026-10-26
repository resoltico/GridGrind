package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.MutationStep;
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
      MutationStep mutationStep,
      FormulaOriginTracker formulaOrigins,
      StepReference authoringStep)
      throws IOException {
    ResolvedStep step =
        new ResolvedStep(
            (ResolvedTarget)
                selectorResolver.resolveMutationTarget(
                    workbook, mutationStep.target(), mutationStep.action()),
            MutationResolution.resolve(mutationStep.action(), bindings));
    FormulaOriginTracker.FormulaWrites writes = formulaOrigins.plannedWrites(workbook, step);
    ResolvedMutationExecutor.apply(workbook, step);
    formulaOrigins.record(workbook, writes, authoringStep);
  }

  void executeStreaming(ExcelStreamingWorkbookWriter writer, MutationStep mutationStep)
      throws IOException {
    writer.apply(
        new ResolvedStep(
            (ResolvedTarget) mutationStep.target(),
            MutationResolution.resolve(mutationStep.action(), bindings)));
  }
}
