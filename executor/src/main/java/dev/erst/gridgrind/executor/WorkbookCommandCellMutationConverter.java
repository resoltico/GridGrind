package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
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
            WorkbookCommandConverter.toExcelCellValue(setCell.value()));
      }
      case CellMutationAction.SetRange setRange -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCellCommand.SetRange(
            selector.sheetName(),
            selector.range(),
            setRange.rows().stream()
                .map(row -> row.stream().map(WorkbookCommandConverter::toExcelCellValue).toList())
                .toList());
      }
      case CellMutationAction.ClearRange _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCellCommand.ClearRange(selector.sheetName(), selector.range());
      }
      case CellMutationAction.SetArrayFormula setArrayFormula -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCellCommand.SetArrayFormula(
            selector.sheetName(),
            selector.range(),
            WorkbookCommandConverter.toExcelArrayFormulaDefinition(setArrayFormula.formula()));
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
            WorkbookCommandConverter.toExcelHyperlink(setHyperlink.target()));
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
            WorkbookCommandConverter.toExcelComment(setComment.comment()));
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
            selector.sheetName(),
            selector.range(),
            WorkbookCommandConverter.toExcelCellStyle(applyStyle.style()));
      }
      case CellMutationAction.AppendRow appendRow ->
          new WorkbookCellCommand.AppendRow(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              appendRow.values().stream().map(WorkbookCommandConverter::toExcelCellValue).toList());
    };
  }
}
