package dev.erst.gridgrind.excel;

import java.util.Arrays;
import java.util.Objects;
import java.util.Set;

/** Applies validated workbook commands to a workbook instance. */
public final class WorkbookCommandExecutor {
  private static final Set<Class<? extends WorkbookCommand>> WORKBOOK_SCOPE_COMMAND_TYPES =
      Set.of(
          WorkbookCommand.CreateSheet.class,
          WorkbookCommand.RenameSheet.class,
          WorkbookCommand.DeleteSheet.class,
          WorkbookCommand.MoveSheet.class,
          WorkbookCommand.CopySheet.class,
          WorkbookCommand.SetActiveSheet.class,
          WorkbookCommand.SetSelectedSheets.class,
          WorkbookCommand.SetSheetVisibility.class,
          WorkbookCommand.SetSheetProtection.class,
          WorkbookCommand.ClearSheetProtection.class,
          WorkbookCommand.SetWorkbookProtection.class,
          WorkbookCommand.ClearWorkbookProtection.class);

  private static final Set<Class<? extends WorkbookCommand>> SHEET_STRUCTURE_COMMAND_TYPES =
      Set.of(
          WorkbookCommand.MergeCells.class,
          WorkbookCommand.UnmergeCells.class,
          WorkbookCommand.SetColumnWidth.class,
          WorkbookCommand.SetRowHeight.class,
          WorkbookCommand.InsertRows.class,
          WorkbookCommand.DeleteRows.class,
          WorkbookCommand.ShiftRows.class,
          WorkbookCommand.InsertColumns.class,
          WorkbookCommand.DeleteColumns.class,
          WorkbookCommand.ShiftColumns.class,
          WorkbookCommand.SetRowVisibility.class,
          WorkbookCommand.SetColumnVisibility.class,
          WorkbookCommand.GroupRows.class,
          WorkbookCommand.UngroupRows.class,
          WorkbookCommand.GroupColumns.class,
          WorkbookCommand.UngroupColumns.class,
          WorkbookCommand.SetSheetPane.class,
          WorkbookCommand.SetSheetZoom.class,
          WorkbookCommand.SetSheetPresentation.class,
          WorkbookCommand.SetPrintLayout.class,
          WorkbookCommand.ClearPrintLayout.class);

  private static final Set<Class<? extends WorkbookCommand>> CELL_VALUE_COMMAND_TYPES =
      Set.of(
          WorkbookCommand.SetCell.class,
          WorkbookCommand.SetRange.class,
          WorkbookCommand.ClearRange.class,
          WorkbookCommand.AppendRow.class,
          WorkbookCommand.AutoSizeColumns.class);

  /** Applies one or more commands in order. */
  public ExcelWorkbook apply(ExcelWorkbook workbook, WorkbookCommand... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return apply(workbook, Arrays.asList(commands));
  }

  /** Applies commands from any iterable source in order. */
  public ExcelWorkbook apply(ExcelWorkbook workbook, Iterable<WorkbookCommand> commands) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    for (WorkbookCommand command : commands) {
      Objects.requireNonNull(command, "command must not be null");
      applyOne(workbook, command);
    }

    return workbook;
  }

  private void applyOne(ExcelWorkbook workbook, WorkbookCommand command) {
    if (isWorkbookScopeCommand(command)) {
      applyWorkbookScopeCommand(workbook, command);
    } else if (isSheetStructureCommand(command)) {
      applySheetStructureCommand(workbook, command);
    } else if (isCellValueCommand(command)) {
      applyCellValueCommand(workbook, command);
    } else {
      applyWorkbookMetadataCommand(workbook, command);
    }
    workbook.markPackageMutated();
    workbook.invalidateFormulaRuntime();
  }

  static boolean isWorkbookScopeCommand(WorkbookCommand command) {
    return WORKBOOK_SCOPE_COMMAND_TYPES.contains(command.getClass());
  }

  static boolean isSheetStructureCommand(WorkbookCommand command) {
    return SHEET_STRUCTURE_COMMAND_TYPES.contains(command.getClass());
  }

  static boolean isCellValueCommand(WorkbookCommand command) {
    return CELL_VALUE_COMMAND_TYPES.contains(command.getClass());
  }

  static void applyWorkbookScopeCommand(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookCommand.CreateSheet createSheet ->
          workbook.getOrCreateSheet(createSheet.sheetName());
      case WorkbookCommand.RenameSheet renameSheet ->
          workbook.renameSheet(renameSheet.sheetName(), renameSheet.newSheetName());
      case WorkbookCommand.DeleteSheet deleteSheet -> workbook.deleteSheet(deleteSheet.sheetName());
      case WorkbookCommand.MoveSheet moveSheet ->
          workbook.moveSheet(moveSheet.sheetName(), moveSheet.targetIndex());
      case WorkbookCommand.CopySheet copySheet ->
          workbook.copySheet(
              copySheet.sourceSheetName(), copySheet.newSheetName(), copySheet.position());
      case WorkbookCommand.SetActiveSheet setActiveSheet ->
          workbook.setActiveSheet(setActiveSheet.sheetName());
      case WorkbookCommand.SetSelectedSheets setSelectedSheets ->
          workbook.setSelectedSheets(setSelectedSheets.sheetNames());
      case WorkbookCommand.SetSheetVisibility setSheetVisibility ->
          workbook.setSheetVisibility(
              setSheetVisibility.sheetName(), setSheetVisibility.visibility());
      case WorkbookCommand.SetSheetProtection setSheetProtection ->
          workbook.setSheetProtection(
              setSheetProtection.sheetName(),
              setSheetProtection.protection(),
              setSheetProtection.password());
      case WorkbookCommand.ClearSheetProtection clearSheetProtection ->
          workbook.clearSheetProtection(clearSheetProtection.sheetName());
      case WorkbookCommand.SetWorkbookProtection setWorkbookProtection ->
          workbook.setWorkbookProtection(setWorkbookProtection.protection());
      case WorkbookCommand.ClearWorkbookProtection _ -> workbook.clearWorkbookProtection();
      default -> throw new IllegalStateException("Unhandled workbook-scope command: " + command);
    }
  }

  static void applySheetStructureCommand(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookCommand.MergeCells mergeCells ->
          workbook.sheet(mergeCells.sheetName()).mergeCells(mergeCells.range());
      case WorkbookCommand.UnmergeCells unmergeCells ->
          workbook.sheet(unmergeCells.sheetName()).unmergeCells(unmergeCells.range());
      case WorkbookCommand.SetColumnWidth setColumnWidth ->
          workbook
              .sheet(setColumnWidth.sheetName())
              .setColumnWidth(
                  setColumnWidth.firstColumnIndex(),
                  setColumnWidth.lastColumnIndex(),
                  setColumnWidth.widthCharacters());
      case WorkbookCommand.SetRowHeight setRowHeight ->
          workbook
              .sheet(setRowHeight.sheetName())
              .setRowHeight(
                  setRowHeight.firstRowIndex(),
                  setRowHeight.lastRowIndex(),
                  setRowHeight.heightPoints());
      case WorkbookCommand.InsertRows insertRows ->
          workbook
              .sheet(insertRows.sheetName())
              .insertRows(insertRows.rowIndex(), insertRows.rowCount());
      case WorkbookCommand.DeleteRows deleteRows ->
          workbook.sheet(deleteRows.sheetName()).deleteRows(deleteRows.rows());
      case WorkbookCommand.ShiftRows shiftRows ->
          workbook.sheet(shiftRows.sheetName()).shiftRows(shiftRows.rows(), shiftRows.delta());
      case WorkbookCommand.InsertColumns insertColumns ->
          workbook
              .sheet(insertColumns.sheetName())
              .insertColumns(insertColumns.columnIndex(), insertColumns.columnCount());
      case WorkbookCommand.DeleteColumns deleteColumns ->
          workbook.sheet(deleteColumns.sheetName()).deleteColumns(deleteColumns.columns());
      case WorkbookCommand.ShiftColumns shiftColumns ->
          workbook
              .sheet(shiftColumns.sheetName())
              .shiftColumns(shiftColumns.columns(), shiftColumns.delta());
      case WorkbookCommand.SetRowVisibility setRowVisibility ->
          workbook
              .sheet(setRowVisibility.sheetName())
              .setRowVisibility(setRowVisibility.rows(), setRowVisibility.hidden());
      case WorkbookCommand.SetColumnVisibility setColumnVisibility ->
          workbook
              .sheet(setColumnVisibility.sheetName())
              .setColumnVisibility(setColumnVisibility.columns(), setColumnVisibility.hidden());
      case WorkbookCommand.GroupRows groupRows ->
          workbook.sheet(groupRows.sheetName()).groupRows(groupRows.rows(), groupRows.collapsed());
      case WorkbookCommand.UngroupRows ungroupRows ->
          workbook.sheet(ungroupRows.sheetName()).ungroupRows(ungroupRows.rows());
      case WorkbookCommand.GroupColumns groupColumns ->
          workbook
              .sheet(groupColumns.sheetName())
              .groupColumns(groupColumns.columns(), groupColumns.collapsed());
      case WorkbookCommand.UngroupColumns ungroupColumns ->
          workbook.sheet(ungroupColumns.sheetName()).ungroupColumns(ungroupColumns.columns());
      case WorkbookCommand.SetSheetPane setSheetPane ->
          workbook.sheet(setSheetPane.sheetName()).setPane(setSheetPane.pane());
      case WorkbookCommand.SetSheetZoom setSheetZoom ->
          workbook.sheet(setSheetZoom.sheetName()).setZoom(setSheetZoom.zoomPercent());
      case WorkbookCommand.SetSheetPresentation setSheetPresentation ->
          workbook
              .sheet(setSheetPresentation.sheetName())
              .setPresentation(setSheetPresentation.presentation());
      case WorkbookCommand.SetPrintLayout setPrintLayout ->
          workbook.sheet(setPrintLayout.sheetName()).setPrintLayout(setPrintLayout.printLayout());
      case WorkbookCommand.ClearPrintLayout clearPrintLayout ->
          workbook.sheet(clearPrintLayout.sheetName()).clearPrintLayout();
      default -> throw new IllegalStateException("Unhandled sheet-structure command: " + command);
    }
  }

  static void applyCellValueCommand(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookCommand.SetCell setCell ->
          workbook.sheet(setCell.sheetName()).setCell(setCell.address(), setCell.value());
      case WorkbookCommand.SetRange setRange ->
          workbook.sheet(setRange.sheetName()).setRange(setRange.range(), setRange.rows());
      case WorkbookCommand.ClearRange clearRange ->
          workbook.sheet(clearRange.sheetName()).clearRange(clearRange.range());
      case WorkbookCommand.AppendRow appendRow ->
          workbook
              .sheet(appendRow.sheetName())
              .appendRow(appendRow.values().toArray(ExcelCellValue[]::new));
      case WorkbookCommand.AutoSizeColumns autoSizeColumns ->
          workbook.sheet(autoSizeColumns.sheetName()).autoSizeColumns();
      default -> throw new IllegalStateException("Unhandled cell-value command: " + command);
    }
  }

  static void applyWorkbookMetadataCommand(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookCommand.SetHyperlink setHyperlink ->
          workbook
              .sheet(setHyperlink.sheetName())
              .setHyperlink(setHyperlink.address(), setHyperlink.target());
      case WorkbookCommand.ClearHyperlink clearHyperlink ->
          workbook.sheet(clearHyperlink.sheetName()).clearHyperlink(clearHyperlink.address());
      case WorkbookCommand.SetComment setComment ->
          workbook
              .sheet(setComment.sheetName())
              .setComment(setComment.address(), setComment.comment());
      case WorkbookCommand.ClearComment clearComment ->
          workbook.sheet(clearComment.sheetName()).clearComment(clearComment.address());
      case WorkbookCommand.SetPicture setPicture ->
          workbook.sheet(setPicture.sheetName()).setPicture(setPicture.picture());
      case WorkbookCommand.SetChart setChart ->
          workbook.sheet(setChart.sheetName()).setChart(setChart.chart());
      case WorkbookCommand.SetShape setShape ->
          workbook.sheet(setShape.sheetName()).setShape(setShape.shape());
      case WorkbookCommand.SetEmbeddedObject setEmbeddedObject ->
          workbook
              .sheet(setEmbeddedObject.sheetName())
              .setEmbeddedObject(setEmbeddedObject.embeddedObject());
      case WorkbookCommand.SetDrawingObjectAnchor setDrawingObjectAnchor ->
          workbook
              .sheet(setDrawingObjectAnchor.sheetName())
              .setDrawingObjectAnchor(
                  setDrawingObjectAnchor.objectName(), setDrawingObjectAnchor.anchor());
      case WorkbookCommand.DeleteDrawingObject deleteDrawingObject ->
          workbook
              .sheet(deleteDrawingObject.sheetName())
              .deleteDrawingObject(deleteDrawingObject.objectName());
      case WorkbookCommand.ApplyStyle applyStyle ->
          workbook.sheet(applyStyle.sheetName()).applyStyle(applyStyle.range(), applyStyle.style());
      case WorkbookCommand.SetDataValidation setDataValidation ->
          workbook
              .sheet(setDataValidation.sheetName())
              .setDataValidation(setDataValidation.range(), setDataValidation.validation());
      case WorkbookCommand.ClearDataValidations clearDataValidations ->
          workbook
              .sheet(clearDataValidations.sheetName())
              .clearDataValidations(clearDataValidations.selection());
      case WorkbookCommand.SetConditionalFormatting setConditionalFormatting ->
          workbook
              .sheet(setConditionalFormatting.sheetName())
              .setConditionalFormatting(setConditionalFormatting.block());
      case WorkbookCommand.ClearConditionalFormatting clearConditionalFormatting ->
          workbook
              .sheet(clearConditionalFormatting.sheetName())
              .clearConditionalFormatting(clearConditionalFormatting.selection());
      case WorkbookCommand.SetAutofilter setAutofilter ->
          workbook
              .sheet(setAutofilter.sheetName())
              .setAutofilter(
                  setAutofilter.range(), setAutofilter.criteria(), setAutofilter.sortState());
      case WorkbookCommand.ClearAutofilter clearAutofilter ->
          workbook.sheet(clearAutofilter.sheetName()).clearAutofilter();
      case WorkbookCommand.SetTable setTable -> workbook.setTable(setTable.definition());
      case WorkbookCommand.SetPivotTable setPivotTable ->
          workbook.setPivotTable(setPivotTable.definition());
      case WorkbookCommand.DeleteTable deleteTable ->
          workbook.deleteTable(deleteTable.name(), deleteTable.sheetName());
      case WorkbookCommand.DeletePivotTable deletePivotTable ->
          workbook.deletePivotTable(deletePivotTable.name(), deletePivotTable.sheetName());
      case WorkbookCommand.SetNamedRange setNamedRange ->
          workbook.setNamedRange(setNamedRange.definition());
      case WorkbookCommand.DeleteNamedRange deleteNamedRange ->
          workbook.deleteNamedRange(deleteNamedRange.name(), deleteNamedRange.scope());
      default -> throw new IllegalStateException("Unhandled workbook-metadata command: " + command);
    }
  }
}
