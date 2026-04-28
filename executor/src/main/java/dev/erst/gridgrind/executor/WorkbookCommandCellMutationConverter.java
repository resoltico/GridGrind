package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.WorkbookCommand;

/** Converts cell-, text-, comment-, and style-oriented mutation families into engine commands. */
final class WorkbookCommandCellMutationConverter {
  private WorkbookCommandCellMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, MutationAction.CellMutationAction action) {
    return switch (action) {
      case MutationAction.SetCell setCell -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCommand.SetCell(
            cellTarget.sheetName(),
            cellTarget.address(),
            WorkbookCommandConverter.toExcelCellValue(setCell.value()));
      }
      case MutationAction.SetRange setRange -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.SetRange(
            selector.sheetName(),
            selector.range(),
            setRange.rows().stream()
                .map(row -> row.stream().map(WorkbookCommandConverter::toExcelCellValue).toList())
                .toList());
      }
      case MutationAction.ClearRange _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.ClearRange(selector.sheetName(), selector.range());
      }
      case MutationAction.SetArrayFormula setArrayFormula -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.SetArrayFormula(
            selector.sheetName(),
            selector.range(),
            WorkbookCommandConverter.toExcelArrayFormulaDefinition(setArrayFormula.formula()));
      }
      case MutationAction.ClearArrayFormula _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCommand.ClearArrayFormula(cellTarget.sheetName(), cellTarget.address());
      }
      case MutationAction.SetHyperlink setHyperlink -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCommand.SetHyperlink(
            cellTarget.sheetName(),
            cellTarget.address(),
            WorkbookCommandConverter.toExcelHyperlink(setHyperlink.target()));
      }
      case MutationAction.ClearHyperlink _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCommand.ClearHyperlink(cellTarget.sheetName(), cellTarget.address());
      }
      case MutationAction.SetComment setComment -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCommand.SetComment(
            cellTarget.sheetName(),
            cellTarget.address(),
            WorkbookCommandConverter.toExcelComment(setComment.comment()));
      }
      case MutationAction.ClearComment _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(
                WorkbookCommandSelectorSupport.cellByAddress(target, action));
        yield new WorkbookCommand.ClearComment(cellTarget.sheetName(), cellTarget.address());
      }
      case MutationAction.ApplyStyle applyStyle -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.ApplyStyle(
            selector.sheetName(),
            selector.range(),
            WorkbookCommandConverter.toExcelCellStyle(applyStyle.style()));
      }
      case MutationAction.AppendRow appendRow ->
          new WorkbookCommand.AppendRow(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              appendRow.values().stream().map(WorkbookCommandConverter::toExcelCellValue).toList());
    };
  }
}
