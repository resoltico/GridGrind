package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.WorkbookCommand;

/**
 * Converts contract mutation steps and style inputs into workbook-core commands.
 *
 * <p>This translation seam intentionally spans the full mutation surface on both sides.
 */
final class WorkbookCommandConverter {
  private WorkbookCommandConverter() {}

  /** Converts one protocol mutation step into the matching workbook-core command. */
  static WorkbookCommand toCommand(MutationStep step) {
    return toCommand(step.target(), step.action());
  }

  /**
   * Converts one protocol mutation action plus selector into the matching workbook-core command.
   */
  static WorkbookCommand toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case WorkbookMutationAction workbookAction ->
          WorkbookCommandWorkbookMutationConverter.toCommand(target, workbookAction);
      case CellMutationAction cellAction ->
          WorkbookCommandCellMutationConverter.toCommand(target, cellAction);
      case DrawingMutationAction drawingAction ->
          WorkbookCommandDrawingMutationConverter.toCommand(target, drawingAction);
      case StructuredMutationAction structuredAction ->
          WorkbookCommandStructuredMutationConverter.toCommand(target, structuredAction);
    };
  }
}
