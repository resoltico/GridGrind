package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.nio.file.Path;
import java.util.List;

/**
 * Converts contract mutation steps and style inputs into workbook-core commands.
 *
 * <p>This translation seam intentionally spans the full mutation surface on both sides.
 */
final class WorkbookCommandConverter {
  private WorkbookCommandConverter() {}

  /** Converts one protocol mutation step into the matching workbook-core command. */
  static WorkbookCommand toCommand(MutationStep step) {
    return toCommand(step, secretUnavailableBindings());
  }

  /** Converts one protocol mutation step using bindings when secret resolution is needed. */
  static WorkbookCommand toCommand(MutationStep step, ExecutionInputBindings bindings) {
    return toCommand(step.target(), step.action(), bindings);
  }

  /**
   * Converts one protocol mutation action plus selector into the matching workbook-core command.
   */
  static WorkbookCommand toCommand(Selector target, MutationAction action) {
    return toCommand(target, action, secretUnavailableBindings());
  }

  private static ExecutionInputBindings secretUnavailableBindings() {
    return new ExecutionInputBindings(
        Path.of("."),
        Path.of("tmp", "gridgrind-command-conversion"),
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum()));
  }

  /**
   * Converts one protocol mutation action plus selector using bindings when secret resolution is
   * needed.
   */
  static WorkbookCommand toCommand(
      Selector target, MutationAction action, ExecutionInputBindings bindings) {
    return switch (action) {
      case WorkbookMutationAction workbookAction ->
          WorkbookCommandWorkbookMutationConverter.toCommand(target, workbookAction, bindings);
      case CellMutationAction cellAction ->
          WorkbookCommandCellMutationConverter.toCommand(target, cellAction);
      case DrawingMutationAction drawingAction ->
          WorkbookCommandDrawingMutationConverter.toCommand(target, drawingAction);
      case StructuredMutationAction structuredAction ->
          WorkbookCommandStructuredMutationConverter.toCommand(target, structuredAction);
    };
  }
}
