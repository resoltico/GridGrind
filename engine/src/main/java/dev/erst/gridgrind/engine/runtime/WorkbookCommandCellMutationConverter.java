package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.excel.WorkbookAnnotationCommand;
import dev.erst.gridgrind.excel.WorkbookCellCommand;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookFormattingCommand;

/** Converts cell-, text-, comment-, and style-oriented mutation families into engine commands. */
final class WorkbookCommandCellMutationConverter {
  private WorkbookCommandCellMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, CellMutationAction action) {
    return switch (action) {
      case CellMutationAction.SetCell setCell -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCellCommand.SetCell(
            cellTarget.sheetName(),
            cellTarget.address(),
            WorkbookCommandCellValueSpecConverter.toResolvedCellWrite(setCell.value()));
      }
      case CellMutationAction.SetRange setRange -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCellCommand.SetRange(
            selector.sheetName().value(),
            selector.range().value(),
            setRange.rows().toCellInputRows().stream()
                .map(
                    row ->
                        row.stream()
                            .map(WorkbookCommandCellValueSpecConverter::toResolvedCellWrite)
                            .toList())
                .toList());
      }
      case CellMutationAction.ClearRange _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCellCommand.ClearRange(
            selector.sheetName().value(), selector.range().value());
      }
      case CellMutationAction.SetArrayFormula setArrayFormula -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCellCommand.SetArrayFormula(
            selector.sheetName().value(),
            selector.range().value(),
            WorkbookCommandCellValueSpecConverter.toResolvedArrayFormula(
                setArrayFormula.formula()));
      }
      case CellMutationAction.ClearArrayFormula _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCellCommand.ClearArrayFormula(
            cellTarget.sheetName(), cellTarget.address());
      }
      case CellMutationAction.SetHyperlink setHyperlink -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookAnnotationCommand.SetHyperlink(
            cellTarget.sheetName(),
            cellTarget.address(),
            WorkbookCommandCellValueSpecConverter.toExcelHyperlink(setHyperlink.target()));
      }
      case CellMutationAction.ClearHyperlink _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookAnnotationCommand.ClearHyperlink(
            cellTarget.sheetName(), cellTarget.address());
      }
      case CellMutationAction.SetComment setComment -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookAnnotationCommand.SetComment(
            cellTarget.sheetName(),
            cellTarget.address(),
            WorkbookCommandCellValueSpecConverter.toResolvedComment(setComment.comment()));
      }
      case CellMutationAction.ClearComment _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookAnnotationCommand.ClearComment(
            cellTarget.sheetName(), cellTarget.address());
      }
      case CellMutationAction.ApplyStyle applyStyle -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookFormattingCommand.ApplyStyle(
            selector.sheetName().value(), selector.range().value(), applyStyle.style());
      }
      case CellMutationAction.AppendRow appendRow ->
          new WorkbookCellCommand.AppendRow(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              appendRow.values().toCellInputs().stream()
                  .map(WorkbookCommandCellValueSpecConverter::toResolvedCellWrite)
                  .toList());
    };
  }
}
