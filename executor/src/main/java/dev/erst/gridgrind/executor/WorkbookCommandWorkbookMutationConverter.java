package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.WorkbookCommand;

/** Converts workbook-, sheet-, and layout-oriented mutation families into engine commands. */
final class WorkbookCommandWorkbookMutationConverter {
  private WorkbookCommandWorkbookMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case MutationAction.EnsureSheet _ ->
          new WorkbookCommand.CreateSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case MutationAction.RenameSheet renameSheet ->
          new WorkbookCommand.RenameSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              renameSheet.newSheetName());
      case MutationAction.DeleteSheet _ ->
          new WorkbookCommand.DeleteSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case MutationAction.MoveSheet moveSheet ->
          new WorkbookCommand.MoveSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              moveSheet.targetIndex());
      case MutationAction.CopySheet copySheet ->
          new WorkbookCommand.CopySheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              copySheet.newSheetName(),
              WorkbookCommandStructuredInputConverter.toExcelSheetCopyPosition(
                  copySheet.position()));
      case MutationAction.SetActiveSheet _ ->
          new WorkbookCommand.SetActiveSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case MutationAction.SetSelectedSheets _ ->
          new WorkbookCommand.SetSelectedSheets(
              WorkbookCommandSelectorSupport.sheetByNames(target, action).names());
      case MutationAction.SetSheetVisibility setSheetVisibility ->
          new WorkbookCommand.SetSheetVisibility(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              setSheetVisibility.visibility());
      case MutationAction.SetSheetProtection setSheetProtection ->
          new WorkbookCommand.SetSheetProtection(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandStructuredInputConverter.toExcelSheetProtectionSettings(
                  setSheetProtection.protection()),
              setSheetProtection.password());
      case MutationAction.ClearSheetProtection _ ->
          new WorkbookCommand.ClearSheetProtection(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case MutationAction.SetWorkbookProtection setWorkbookProtection ->
          new WorkbookCommand.SetWorkbookProtection(
              WorkbookCommandStructuredInputConverter.toExcelWorkbookProtectionSettings(
                  setWorkbookProtection.protection()));
      case MutationAction.ClearWorkbookProtection _ ->
          new WorkbookCommand.ClearWorkbookProtection();
      case MutationAction.MergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.MergeCells(selector.sheetName(), selector.range());
      }
      case MutationAction.UnmergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookCommand.UnmergeCells(selector.sheetName(), selector.range());
      }
      case MutationAction.SetColumnWidth setColumnWidth -> {
        ColumnBandSelector.Span selector =
            WorkbookCommandSelectorSupport.columnSpan(target, action);
        yield new WorkbookCommand.SetColumnWidth(
            selector.sheetName(),
            selector.firstColumnIndex(),
            selector.lastColumnIndex(),
            setColumnWidth.widthCharacters());
      }
      case MutationAction.SetRowHeight setRowHeight -> {
        RowBandSelector.Span selector = WorkbookCommandSelectorSupport.rowSpan(target, action);
        yield new WorkbookCommand.SetRowHeight(
            selector.sheetName(),
            selector.firstRowIndex(),
            selector.lastRowIndex(),
            setRowHeight.heightPoints());
      }
      case MutationAction.InsertRows _ -> {
        RowBandSelector.Insertion selector =
            WorkbookCommandSelectorSupport.rowInsertion(target, action);
        yield new WorkbookCommand.InsertRows(
            selector.sheetName(), selector.beforeRowIndex(), selector.rowCount());
      }
      case MutationAction.DeleteRows _ ->
          new WorkbookCommand.DeleteRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)));
      case MutationAction.ShiftRows shiftRows ->
          new WorkbookCommand.ShiftRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)),
              shiftRows.delta());
      case MutationAction.InsertColumns _ -> {
        ColumnBandSelector.Insertion selector =
            WorkbookCommandSelectorSupport.columnInsertion(target, action);
        yield new WorkbookCommand.InsertColumns(
            selector.sheetName(), selector.beforeColumnIndex(), selector.columnCount());
      }
      case MutationAction.DeleteColumns _ ->
          new WorkbookCommand.DeleteColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)));
      case MutationAction.ShiftColumns shiftColumns ->
          new WorkbookCommand.ShiftColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              shiftColumns.delta());
      case MutationAction.SetRowVisibility setRowVisibility ->
          new WorkbookCommand.SetRowVisibility(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)),
              setRowVisibility.hidden());
      case MutationAction.SetColumnVisibility setColumnVisibility ->
          new WorkbookCommand.SetColumnVisibility(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              setColumnVisibility.hidden());
      case MutationAction.GroupRows groupRows ->
          new WorkbookCommand.GroupRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)),
              groupRows.collapsed());
      case MutationAction.UngroupRows _ ->
          new WorkbookCommand.UngroupRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)));
      case MutationAction.GroupColumns groupColumns ->
          new WorkbookCommand.GroupColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              groupColumns.collapsed());
      case MutationAction.UngroupColumns _ ->
          new WorkbookCommand.UngroupColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)));
      case MutationAction.SetSheetPane setSheetPane ->
          new WorkbookCommand.SetSheetPane(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandStructuredInputConverter.toExcelSheetPane(setSheetPane.pane()));
      case MutationAction.SetSheetZoom setSheetZoom ->
          new WorkbookCommand.SetSheetZoom(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              setSheetZoom.zoomPercent());
      case MutationAction.SetSheetPresentation setSheetPresentation ->
          new WorkbookCommand.SetSheetPresentation(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandStructuredInputConverter.toExcelSheetPresentation(
                  setSheetPresentation.presentation()));
      case MutationAction.SetPrintLayout setPrintLayout ->
          new WorkbookCommand.SetPrintLayout(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandStructuredInputConverter.toExcelPrintLayout(
                  setPrintLayout.printLayout()));
      case MutationAction.ClearPrintLayout _ ->
          new WorkbookCommand.ClearPrintLayout(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case MutationAction.AutoSizeColumns _ ->
          new WorkbookCommand.AutoSizeColumns(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      default -> null;
    };
  }
}
