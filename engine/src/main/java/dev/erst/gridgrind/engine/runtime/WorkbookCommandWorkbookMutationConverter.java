package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookLayoutCommand;
import dev.erst.gridgrind.excel.WorkbookSheetCommand;
import dev.erst.gridgrind.excel.WorkbookStructureCommand;
import java.util.Optional;

/** Converts workbook-, sheet-, and layout-oriented mutation families into engine commands. */
final class WorkbookCommandWorkbookMutationConverter {
  private WorkbookCommandWorkbookMutationConverter() {}

  static WorkbookCommand toCommand(
      Selector target, WorkbookMutationAction action, ExecutionInputBindings bindings) {
    return switch (action) {
      case WorkbookMutationAction.EnsureSheet _ ->
          new WorkbookSheetCommand.CreateSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value());
      case WorkbookMutationAction.RenameSheet renameSheet ->
          new WorkbookSheetCommand.RenameSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              renameSheet.newSheetName().value());
      case WorkbookMutationAction.DeleteSheet _ ->
          new WorkbookSheetCommand.DeleteSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value());
      case WorkbookMutationAction.MoveSheet moveSheet ->
          new WorkbookSheetCommand.MoveSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              moveSheet.targetIndex());
      case WorkbookMutationAction.CopySheet copySheet ->
          new WorkbookSheetCommand.CopySheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              copySheet.newSheetName().value(),
              WorkbookCommandLayoutInputConverter.toExcelSheetCopyPosition(copySheet.position()));
      case WorkbookMutationAction.SetActiveSheet _ ->
          new WorkbookSheetCommand.SetActiveSheet(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value());
      case WorkbookMutationAction.SetSelectedSheets _ ->
          new WorkbookSheetCommand.SetSelectedSheets(
              WorkbookCommandSelectorSupport.sheetByNames(target, action).names().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                  .toList());
      case WorkbookMutationAction.SetSheetVisibility setSheetVisibility ->
          new WorkbookSheetCommand.SetSheetVisibility(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              setSheetVisibility.visibility());
      case WorkbookMutationAction.SetSheetProtection setSheetProtection ->
          new WorkbookSheetCommand.SetSheetProtection(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              WorkbookCommandLayoutInputConverter.toExcelSheetProtectionSettings(
                  setSheetProtection.protection()),
              resolveOptionalSecret(setSheetProtection.passwordRef(), bindings));
      case WorkbookMutationAction.ClearSheetProtection _ ->
          new WorkbookSheetCommand.ClearSheetProtection(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value());
      case WorkbookMutationAction.SetWorkbookProtection setWorkbookProtection ->
          new WorkbookSheetCommand.SetWorkbookProtection(
              WorkbookCommandLayoutInputConverter.toExcelWorkbookProtectionSettings(
                  setWorkbookProtection.protection(), bindings));
      case WorkbookMutationAction.ClearWorkbookProtection _ ->
          new WorkbookSheetCommand.ClearWorkbookProtection();
      case WorkbookMutationAction.MergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookStructureCommand.MergeCells(
            selector.sheetName().value(), selector.range().value());
      }
      case WorkbookMutationAction.UnmergeCells _ -> {
        RangeSelector.ByRange selector =
            WorkbookCommandSelectorSupport.rangeByRange(target, action);
        yield new WorkbookStructureCommand.UnmergeCells(
            selector.sheetName().value(), selector.range().value());
      }
      case WorkbookMutationAction.SetColumnWidth setColumnWidth -> {
        ColumnBandSelector.Span selector =
            WorkbookCommandSelectorSupport.columnSpan(target, action);
        yield new WorkbookStructureCommand.SetColumnWidth(
            selector.sheetName().value(),
            selector.firstColumnIndex(),
            selector.lastColumnIndex(),
            setColumnWidth.widthCharacters());
      }
      case WorkbookMutationAction.SetRowHeight setRowHeight -> {
        RowBandSelector.Span selector = WorkbookCommandSelectorSupport.rowSpan(target, action);
        yield new WorkbookStructureCommand.SetRowHeight(
            selector.sheetName().value(),
            selector.firstRowIndex(),
            selector.lastRowIndex(),
            setRowHeight.heightPoints());
      }
      case WorkbookMutationAction.InsertRows _ -> {
        RowBandSelector.Insertion selector =
            WorkbookCommandSelectorSupport.rowInsertion(target, action);
        yield new WorkbookStructureCommand.InsertRows(
            selector.sheetName().value(), selector.beforeRowIndex(), selector.rowCount());
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
            selector.sheetName().value(), selector.beforeColumnIndex(), selector.columnCount());
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
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              WorkbookCommandLayoutInputConverter.toExcelSheetPane(setSheetPane.pane()));
      case WorkbookMutationAction.SetSheetZoom setSheetZoom ->
          new WorkbookLayoutCommand.SetSheetZoom(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              setSheetZoom.zoomPercent());
      case WorkbookMutationAction.SetSheetPresentation setSheetPresentation ->
          new WorkbookLayoutCommand.SetSheetPresentation(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              WorkbookCommandPresentationInputConverter.toExcelSheetPresentation(
                  setSheetPresentation.presentation()));
      case WorkbookMutationAction.SetPrintLayout setPrintLayout ->
          new WorkbookLayoutCommand.SetPrintLayout(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value(),
              WorkbookCommandLayoutInputConverter.toResolvedPrintLayout(
                  setPrintLayout.printLayout()));
      case WorkbookMutationAction.ClearPrintLayout _ ->
          new WorkbookLayoutCommand.ClearPrintLayout(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value());
      case WorkbookMutationAction.AutoSizeColumns _ ->
          new WorkbookLayoutCommand.AutoSizeColumns(
              WorkbookCommandSelectorSupport.sheetByName(target, action).name().value());
    };
  }

  static WorkbookCommand toCommand(Selector target, WorkbookMutationAction action) {
    return WorkbookCommandConverter.toCommand(target, action);
  }

  private static Optional<String> resolveOptionalSecret(
      Optional<dev.erst.gridgrind.core.source.SecretReference> reference,
      ExecutionInputBindings bindings) {
    if (reference.isEmpty()) {
      return Optional.empty();
    }
    if (bindings == null) {
      throw new IllegalStateException("secret-bearing workbook actions require execution bindings");
    }
    return Optional.of(bindings.resolveSecret(reference.orElseThrow()));
  }
}
