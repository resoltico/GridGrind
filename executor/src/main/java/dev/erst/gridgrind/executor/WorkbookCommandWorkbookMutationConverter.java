package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookLayoutCommand;
import dev.erst.gridgrind.excel.WorkbookSheetCommand;
import dev.erst.gridgrind.excel.WorkbookStructureCommand;

/** Converts workbook-, sheet-, and layout-oriented mutation families into engine commands. */
final class WorkbookCommandWorkbookMutationConverter {
  private WorkbookCommandWorkbookMutationConverter() {}

  static WorkbookCommand toCommand(Selector target, WorkbookMutationAction action) {
    return switch (action) {
      case WorkbookMutationAction.EnsureSheet _ ->
          new WorkbookSheetCommand.CreateSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case WorkbookMutationAction.RenameSheet renameSheet ->
          new WorkbookSheetCommand.RenameSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              renameSheet.newSheetName());
      case WorkbookMutationAction.DeleteSheet _ ->
          new WorkbookSheetCommand.DeleteSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case WorkbookMutationAction.MoveSheet moveSheet ->
          new WorkbookSheetCommand.MoveSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              moveSheet.targetIndex());
      case WorkbookMutationAction.CopySheet copySheet ->
          new WorkbookSheetCommand.CopySheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              copySheet.newSheetName(),
              WorkbookCommandLayoutInputConverter.toExcelSheetCopyPosition(copySheet.position()));
      case WorkbookMutationAction.SetActiveSheet _ ->
          new WorkbookSheetCommand.SetActiveSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case WorkbookMutationAction.SetSelectedSheets _ ->
          new WorkbookSheetCommand.SetSelectedSheets(
              WorkbookCommandSelectorSupport.sheetByNames(target, action).names());
      case WorkbookMutationAction.SetSheetVisibility setSheetVisibility ->
          new WorkbookSheetCommand.SetSheetVisibility(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              setSheetVisibility.visibility());
      case WorkbookMutationAction.SetSheetProtection setSheetProtection ->
          new WorkbookSheetCommand.SetSheetProtection(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandLayoutInputConverter.toExcelSheetProtectionSettings(
                  setSheetProtection.protection()),
              setSheetProtection.password());
      case WorkbookMutationAction.ClearSheetProtection _ ->
          new WorkbookSheetCommand.ClearSheetProtection(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case WorkbookMutationAction.SetWorkbookProtection setWorkbookProtection ->
          new WorkbookSheetCommand.SetWorkbookProtection(
              WorkbookCommandLayoutInputConverter.toExcelWorkbookProtectionSettings(
                  setWorkbookProtection.protection()));
      case WorkbookMutationAction.ClearWorkbookProtection _ ->
          new WorkbookSheetCommand.ClearWorkbookProtection();
      case WorkbookMutationAction.MergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookStructureCommand.MergeCells(selector.sheetName(), selector.range());
      }
      case WorkbookMutationAction.UnmergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookStructureCommand.UnmergeCells(selector.sheetName(), selector.range());
      }
      case WorkbookMutationAction.SetColumnWidth setColumnWidth -> {
        ColumnBandSelector.Span selector =
            WorkbookCommandSelectorSupport.columnSpan(target, action);
        yield new WorkbookStructureCommand.SetColumnWidth(
            selector.sheetName(),
            selector.firstColumnIndex(),
            selector.lastColumnIndex(),
            setColumnWidth.widthCharacters());
      }
      case WorkbookMutationAction.SetRowHeight setRowHeight -> {
        RowBandSelector.Span selector = WorkbookCommandSelectorSupport.rowSpan(target, action);
        yield new WorkbookStructureCommand.SetRowHeight(
            selector.sheetName(),
            selector.firstRowIndex(),
            selector.lastRowIndex(),
            setRowHeight.heightPoints());
      }
      case WorkbookMutationAction.InsertRows _ -> {
        RowBandSelector.Insertion selector =
            WorkbookCommandSelectorSupport.rowInsertion(target, action);
        yield new WorkbookStructureCommand.InsertRows(
            selector.sheetName(), selector.beforeRowIndex(), selector.rowCount());
      }
      case WorkbookMutationAction.DeleteRows _ ->
          new WorkbookStructureCommand.DeleteRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)));
      case WorkbookMutationAction.ShiftRows shiftRows ->
          new WorkbookStructureCommand.ShiftRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)),
              shiftRows.delta());
      case WorkbookMutationAction.InsertColumns _ -> {
        ColumnBandSelector.Insertion selector =
            WorkbookCommandSelectorSupport.columnInsertion(target, action);
        yield new WorkbookStructureCommand.InsertColumns(
            selector.sheetName(), selector.beforeColumnIndex(), selector.columnCount());
      }
      case WorkbookMutationAction.DeleteColumns _ ->
          new WorkbookStructureCommand.DeleteColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)));
      case WorkbookMutationAction.ShiftColumns shiftColumns ->
          new WorkbookStructureCommand.ShiftColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              shiftColumns.delta());
      case WorkbookMutationAction.SetRowVisibility setRowVisibility ->
          new WorkbookStructureCommand.SetRowVisibility(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)),
              setRowVisibility.hidden());
      case WorkbookMutationAction.SetColumnVisibility setColumnVisibility ->
          new WorkbookStructureCommand.SetColumnVisibility(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              setColumnVisibility.hidden());
      case WorkbookMutationAction.GroupRows groupRows ->
          new WorkbookStructureCommand.GroupRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)),
              groupRows.collapsed());
      case WorkbookMutationAction.UngroupRows _ ->
          new WorkbookStructureCommand.UngroupRows(
              SelectorConverter.toSheetName(WorkbookCommandSelectorSupport.rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(
                  WorkbookCommandSelectorSupport.rowSpan(target, action)));
      case WorkbookMutationAction.GroupColumns groupColumns ->
          new WorkbookStructureCommand.GroupColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              groupColumns.collapsed());
      case WorkbookMutationAction.UngroupColumns _ ->
          new WorkbookStructureCommand.UngroupColumns(
              SelectorConverter.toSheetName(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(
                  WorkbookCommandSelectorSupport.columnSpan(target, action)));
      case WorkbookMutationAction.SetSheetPane setSheetPane ->
          new WorkbookLayoutCommand.SetSheetPane(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandLayoutInputConverter.toExcelSheetPane(setSheetPane.pane()));
      case WorkbookMutationAction.SetSheetZoom setSheetZoom ->
          new WorkbookLayoutCommand.SetSheetZoom(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              setSheetZoom.zoomPercent());
      case WorkbookMutationAction.SetSheetPresentation setSheetPresentation ->
          new WorkbookLayoutCommand.SetSheetPresentation(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandLayoutInputConverter.toExcelSheetPresentation(
                  setSheetPresentation.presentation()));
      case WorkbookMutationAction.SetPrintLayout setPrintLayout ->
          new WorkbookLayoutCommand.SetPrintLayout(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
              WorkbookCommandLayoutInputConverter.toExcelPrintLayout(setPrintLayout.printLayout()));
      case WorkbookMutationAction.ClearPrintLayout _ ->
          new WorkbookLayoutCommand.ClearPrintLayout(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
      case WorkbookMutationAction.AutoSizeColumns _ ->
          new WorkbookLayoutCommand.AutoSizeColumns(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name());
    };
  }
}
