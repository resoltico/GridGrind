package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter;
import java.io.IOException;
import java.util.Objects;

/**
 * Resolves one mutation, validates normal formulas in workbook state, then records their origin.
 */
final class FormulaAwareMutationStepExecutor {
  private final WorkbookExecutionEngine workbookEngine;
  private final SemanticSelectorResolver selectorResolver;

  FormulaAwareMutationStepExecutor(
      WorkbookExecutionEngine workbookEngine, SemanticSelectorResolver selectorResolver) {
    this.workbookEngine = Objects.requireNonNull(workbookEngine, "workbookEngine must not be null");
    this.selectorResolver =
        Objects.requireNonNull(selectorResolver, "selectorResolver must not be null");
  }

  void execute(
      ExcelWorkbook workbook,
      MutationStep mutationStep,
      FormulaOriginTracker formulaOrigins,
      StepReference authoringStep)
      throws IOException {
    WorkbookCommand command =
        WorkbookCommandConverter.toCommand(
            selectorResolver.resolveMutationTarget(
                workbook, mutationStep.target(), mutationStep.action()),
            mutationStep.action());
    FormulaOriginTracker.FormulaWrites writes = formulaOrigins.plannedWrites(workbook, command);
    workbookEngine.apply(workbook, command);
    formulaOrigins.record(workbook, writes, authoringStep);
  }

  void executeStreaming(ExcelStreamingWorkbookWriter writer, MutationStep mutationStep)
      throws IOException {
    writer.apply(WorkbookCommandConverter.toCommand(mutationStep));
  }
}
