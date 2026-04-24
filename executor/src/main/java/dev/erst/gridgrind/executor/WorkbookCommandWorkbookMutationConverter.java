package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.util.Optional;

/** Converts workbook-, sheet-, and layout-oriented mutation families into engine commands. */
final class WorkbookCommandWorkbookMutationConverter {
  private WorkbookCommandWorkbookMutationConverter() {}

  static Optional<WorkbookCommand> toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case MutationAction.EnsureSheet _ ->
          Optional.of(
              new WorkbookCommand.CreateSheet(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name()));
      case MutationAction.RenameSheet renameSheet ->
          Optional.of(
              new WorkbookCommand.RenameSheet(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  renameSheet.newSheetName()));
      case MutationAction.DeleteSheet _ ->
          Optional.of(
              new WorkbookCommand.DeleteSheet(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name()));
      case MutationAction.MoveSheet moveSheet ->
          Optional.of(
              new WorkbookCommand.MoveSheet(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  moveSheet.targetIndex()));
      case MutationAction.CopySheet copySheet ->
          Optional.of(
              new WorkbookCommand.CopySheet(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  copySheet.newSheetName(),
                  WorkbookCommandStructuredInputConverter.toExcelSheetCopyPosition(
                      copySheet.position())));
      case MutationAction.SetActiveSheet _ ->
          Optional.of(
              new WorkbookCommand.SetActiveSheet(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name()));
      case MutationAction.SetSelectedSheets _ ->
          Optional.of(
              new WorkbookCommand.SetSelectedSheets(
                  WorkbookCommandSelectorSupport.sheetByNames(target, action).names()));
      case MutationAction.SetSheetVisibility setSheetVisibility ->
          Optional.of(
              new WorkbookCommand.SetSheetVisibility(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  setSheetVisibility.visibility()));
      case MutationAction.SetSheetProtection setSheetProtection ->
          Optional.of(
              new WorkbookCommand.SetSheetProtection(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  WorkbookCommandStructuredInputConverter.toExcelSheetProtectionSettings(
                      setSheetProtection.protection()),
                  setSheetProtection.password()));
      case MutationAction.ClearSheetProtection _ ->
          Optional.of(
              new WorkbookCommand.ClearSheetProtection(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name()));
      case MutationAction.SetWorkbookProtection setWorkbookProtection ->
          Optional.of(
              new WorkbookCommand.SetWorkbookProtection(
                  WorkbookCommandStructuredInputConverter.toExcelWorkbookProtectionSettings(
                      setWorkbookProtection.protection())));
      case MutationAction.ClearWorkbookProtection _ ->
          Optional.of(new WorkbookCommand.ClearWorkbookProtection());
      case MutationAction.MergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield Optional.of(new WorkbookCommand.MergeCells(selector.sheetName(), selector.range()));
      }
      case MutationAction.UnmergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield Optional.of(new WorkbookCommand.UnmergeCells(selector.sheetName(), selector.range()));
      }
      case MutationAction.SetColumnWidth setColumnWidth -> {
        ColumnBandSelector.Span selector =
            WorkbookCommandSelectorSupport.columnSpan(target, action);
        yield Optional.of(
            new WorkbookCommand.SetColumnWidth(
                selector.sheetName(),
                selector.firstColumnIndex(),
                selector.lastColumnIndex(),
                setColumnWidth.widthCharacters()));
      }
      case MutationAction.SetRowHeight setRowHeight -> {
        RowBandSelector.Span selector = WorkbookCommandSelectorSupport.rowSpan(target, action);
        yield Optional.of(
            new WorkbookCommand.SetRowHeight(
                selector.sheetName(),
                selector.firstRowIndex(),
                selector.lastRowIndex(),
                setRowHeight.heightPoints()));
      }
      case MutationAction.InsertRows _ -> {
        RowBandSelector.Insertion selector =
            WorkbookCommandSelectorSupport.rowInsertion(target, action);
        yield Optional.of(
            new WorkbookCommand.InsertRows(
                selector.sheetName(), selector.beforeRowIndex(), selector.rowCount()));
      }
      case MutationAction.DeleteRows _ ->
          Optional.of(
              new WorkbookCommand.DeleteRows(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  SelectorConverter.toExcelRowSpan(
                      WorkbookCommandSelectorSupport.rowSpan(target, action))));
      case MutationAction.ShiftRows shiftRows ->
          Optional.of(
              new WorkbookCommand.ShiftRows(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  SelectorConverter.toExcelRowSpan(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  shiftRows.delta()));
      case MutationAction.InsertColumns _ -> {
        ColumnBandSelector.Insertion selector =
            WorkbookCommandSelectorSupport.columnInsertion(target, action);
        yield Optional.of(
            new WorkbookCommand.InsertColumns(
                selector.sheetName(), selector.beforeColumnIndex(), selector.columnCount()));
      }
      case MutationAction.DeleteColumns _ ->
          Optional.of(
              new WorkbookCommand.DeleteColumns(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  SelectorConverter.toExcelColumnSpan(
                      WorkbookCommandSelectorSupport.columnSpan(target, action))));
      case MutationAction.ShiftColumns shiftColumns ->
          Optional.of(
              new WorkbookCommand.ShiftColumns(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  SelectorConverter.toExcelColumnSpan(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  shiftColumns.delta()));
      case MutationAction.SetRowVisibility setRowVisibility ->
          Optional.of(
              new WorkbookCommand.SetRowVisibility(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  SelectorConverter.toExcelRowSpan(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  setRowVisibility.hidden()));
      case MutationAction.SetColumnVisibility setColumnVisibility ->
          Optional.of(
              new WorkbookCommand.SetColumnVisibility(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  SelectorConverter.toExcelColumnSpan(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  setColumnVisibility.hidden()));
      case MutationAction.GroupRows groupRows ->
          Optional.of(
              new WorkbookCommand.GroupRows(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  SelectorConverter.toExcelRowSpan(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  groupRows.collapsed()));
      case MutationAction.UngroupRows _ ->
          Optional.of(
              new WorkbookCommand.UngroupRows(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.rowSpan(target, action)),
                  SelectorConverter.toExcelRowSpan(
                      WorkbookCommandSelectorSupport.rowSpan(target, action))));
      case MutationAction.GroupColumns groupColumns ->
          Optional.of(
              new WorkbookCommand.GroupColumns(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  SelectorConverter.toExcelColumnSpan(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  groupColumns.collapsed()));
      case MutationAction.UngroupColumns _ ->
          Optional.of(
              new WorkbookCommand.UngroupColumns(
                  SelectorConverter.toSheetName(
                      WorkbookCommandSelectorSupport.columnSpan(target, action)),
                  SelectorConverter.toExcelColumnSpan(
                      WorkbookCommandSelectorSupport.columnSpan(target, action))));
      case MutationAction.SetSheetPane setSheetPane ->
          Optional.of(
              new WorkbookCommand.SetSheetPane(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  WorkbookCommandStructuredInputConverter.toExcelSheetPane(setSheetPane.pane())));
      case MutationAction.SetSheetZoom setSheetZoom ->
          Optional.of(
              new WorkbookCommand.SetSheetZoom(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  setSheetZoom.zoomPercent()));
      case MutationAction.SetSheetPresentation setSheetPresentation ->
          Optional.of(
              new WorkbookCommand.SetSheetPresentation(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  WorkbookCommandStructuredInputConverter.toExcelSheetPresentation(
                      setSheetPresentation.presentation())));
      case MutationAction.SetPrintLayout setPrintLayout ->
          Optional.of(
              new WorkbookCommand.SetPrintLayout(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name(),
                  WorkbookCommandStructuredInputConverter.toExcelPrintLayout(
                      setPrintLayout.printLayout())));
      case MutationAction.ClearPrintLayout _ ->
          Optional.of(
              new WorkbookCommand.ClearPrintLayout(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name()));
      case MutationAction.AutoSizeColumns _ ->
          Optional.of(
              new WorkbookCommand.AutoSizeColumns(
                  WorkbookCommandSelectorSupport.sheetByName(target, action).name()));
      default -> Optional.empty();
    };
  }
}
