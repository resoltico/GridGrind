package dev.erst.gridgrind.excel;

import java.util.Arrays;
import java.util.Objects;
import java.util.Set;

/** Applies validated workbook commands to a workbook instance. */
public final class WorkbookCommandExecutor {
  private static final Set<Class<? extends WorkbookCommand>> WORKBOOK_SCOPE_COMMAND_TYPES =
      Set.of(
          WorkbookSheetCommand.CreateSheet.class,
          WorkbookSheetCommand.RenameSheet.class,
          WorkbookSheetCommand.DeleteSheet.class,
          WorkbookSheetCommand.MoveSheet.class,
          WorkbookSheetCommand.CopySheet.class,
          WorkbookSheetCommand.SetActiveSheet.class,
          WorkbookSheetCommand.SetSelectedSheets.class,
          WorkbookSheetCommand.SetSheetVisibility.class,
          WorkbookSheetCommand.SetSheetProtection.class,
          WorkbookSheetCommand.ClearSheetProtection.class,
          WorkbookSheetCommand.SetWorkbookProtection.class,
          WorkbookSheetCommand.ClearWorkbookProtection.class);

  private static final Set<Class<? extends WorkbookCommand>> SHEET_STRUCTURE_COMMAND_TYPES =
      Set.of(
          WorkbookStructureCommand.MergeCells.class,
          WorkbookStructureCommand.UnmergeCells.class,
          WorkbookStructureCommand.SetColumnWidth.class,
          WorkbookStructureCommand.SetRowHeight.class,
          WorkbookStructureCommand.InsertRows.class,
          WorkbookStructureCommand.DeleteRows.class,
          WorkbookStructureCommand.ShiftRows.class,
          WorkbookStructureCommand.InsertColumns.class,
          WorkbookStructureCommand.DeleteColumns.class,
          WorkbookStructureCommand.ShiftColumns.class,
          WorkbookStructureCommand.SetRowVisibility.class,
          WorkbookStructureCommand.SetColumnVisibility.class,
          WorkbookStructureCommand.GroupRows.class,
          WorkbookStructureCommand.UngroupRows.class,
          WorkbookStructureCommand.GroupColumns.class,
          WorkbookStructureCommand.UngroupColumns.class,
          WorkbookLayoutCommand.SetSheetPane.class,
          WorkbookLayoutCommand.SetSheetZoom.class,
          WorkbookLayoutCommand.SetSheetPresentation.class,
          WorkbookLayoutCommand.SetPrintLayout.class,
          WorkbookLayoutCommand.ClearPrintLayout.class);

  private static final Set<Class<? extends WorkbookCommand>> CELL_VALUE_COMMAND_TYPES =
      Set.of(
          WorkbookCellCommand.SetCell.class,
          WorkbookCellCommand.SetRange.class,
          WorkbookCellCommand.ClearRange.class,
          WorkbookCellCommand.SetArrayFormula.class,
          WorkbookCellCommand.ClearArrayFormula.class,
          WorkbookCellCommand.AppendRow.class,
          WorkbookLayoutCommand.AutoSizeColumns.class);

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
      case WorkbookSheetCommand.CreateSheet createSheet ->
          workbook.getOrCreateSheet(createSheet.sheetName());
      case WorkbookSheetCommand.RenameSheet renameSheet ->
          workbook.renameSheet(renameSheet.sheetName(), renameSheet.newSheetName());
      case WorkbookSheetCommand.DeleteSheet deleteSheet ->
          workbook.deleteSheet(deleteSheet.sheetName());
      case WorkbookSheetCommand.MoveSheet moveSheet ->
          workbook.moveSheet(moveSheet.sheetName(), moveSheet.targetIndex());
      case WorkbookSheetCommand.CopySheet copySheet ->
          workbook.copySheet(
              copySheet.sourceSheetName(), copySheet.newSheetName(), copySheet.position());
      case WorkbookSheetCommand.SetActiveSheet setActiveSheet ->
          workbook.setActiveSheet(setActiveSheet.sheetName());
      case WorkbookSheetCommand.SetSelectedSheets setSelectedSheets ->
          workbook.setSelectedSheets(setSelectedSheets.sheetNames());
      case WorkbookSheetCommand.SetSheetVisibility setSheetVisibility ->
          workbook.setSheetVisibility(
              setSheetVisibility.sheetName(), setSheetVisibility.visibility());
      case WorkbookSheetCommand.SetSheetProtection setSheetProtection ->
          workbook.setSheetProtection(
              setSheetProtection.sheetName(),
              setSheetProtection.protection(),
              setSheetProtection.password());
      case WorkbookSheetCommand.ClearSheetProtection clearSheetProtection ->
          workbook.clearSheetProtection(clearSheetProtection.sheetName());
      case WorkbookSheetCommand.SetWorkbookProtection setWorkbookProtection ->
          workbook.setWorkbookProtection(setWorkbookProtection.protection());
      case WorkbookSheetCommand.ClearWorkbookProtection _ -> workbook.clearWorkbookProtection();
      default -> throw new IllegalStateException("Unhandled workbook-scope command: " + command);
    }
  }

  static void applySheetStructureCommand(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookStructureCommand.MergeCells mergeCells ->
          workbook.sheet(mergeCells.sheetName()).mergeCells(mergeCells.range());
      case WorkbookStructureCommand.UnmergeCells unmergeCells ->
          workbook.sheet(unmergeCells.sheetName()).unmergeCells(unmergeCells.range());
      case WorkbookStructureCommand.SetColumnWidth setColumnWidth ->
          workbook
              .sheet(setColumnWidth.sheetName())
              .setColumnWidth(
                  setColumnWidth.firstColumnIndex(),
                  setColumnWidth.lastColumnIndex(),
                  setColumnWidth.widthCharacters());
      case WorkbookStructureCommand.SetRowHeight setRowHeight ->
          workbook
              .sheet(setRowHeight.sheetName())
              .setRowHeight(
                  setRowHeight.firstRowIndex(),
                  setRowHeight.lastRowIndex(),
                  setRowHeight.heightPoints());
      case WorkbookStructureCommand.InsertRows insertRows ->
          workbook
              .sheet(insertRows.sheetName())
              .insertRows(insertRows.rowIndex(), insertRows.rowCount());
      case WorkbookStructureCommand.DeleteRows deleteRows ->
          workbook.sheet(deleteRows.sheetName()).deleteRows(deleteRows.rows());
      case WorkbookStructureCommand.ShiftRows shiftRows ->
          workbook.sheet(shiftRows.sheetName()).shiftRows(shiftRows.rows(), shiftRows.delta());
      case WorkbookStructureCommand.InsertColumns insertColumns ->
          workbook
              .sheet(insertColumns.sheetName())
              .insertColumns(insertColumns.columnIndex(), insertColumns.columnCount());
      case WorkbookStructureCommand.DeleteColumns deleteColumns ->
          workbook.sheet(deleteColumns.sheetName()).deleteColumns(deleteColumns.columns());
      case WorkbookStructureCommand.ShiftColumns shiftColumns ->
          workbook
              .sheet(shiftColumns.sheetName())
              .shiftColumns(shiftColumns.columns(), shiftColumns.delta());
      case WorkbookStructureCommand.SetRowVisibility setRowVisibility ->
          workbook
              .sheet(setRowVisibility.sheetName())
              .setRowVisibility(setRowVisibility.rows(), setRowVisibility.hidden());
      case WorkbookStructureCommand.SetColumnVisibility setColumnVisibility ->
          workbook
              .sheet(setColumnVisibility.sheetName())
              .setColumnVisibility(setColumnVisibility.columns(), setColumnVisibility.hidden());
      case WorkbookStructureCommand.GroupRows groupRows ->
          workbook.sheet(groupRows.sheetName()).groupRows(groupRows.rows(), groupRows.collapsed());
      case WorkbookStructureCommand.UngroupRows ungroupRows ->
          workbook.sheet(ungroupRows.sheetName()).ungroupRows(ungroupRows.rows());
      case WorkbookStructureCommand.GroupColumns groupColumns ->
          workbook
              .sheet(groupColumns.sheetName())
              .groupColumns(groupColumns.columns(), groupColumns.collapsed());
      case WorkbookStructureCommand.UngroupColumns ungroupColumns ->
          workbook.sheet(ungroupColumns.sheetName()).ungroupColumns(ungroupColumns.columns());
      case WorkbookLayoutCommand.SetSheetPane setSheetPane ->
          workbook.sheet(setSheetPane.sheetName()).setPane(setSheetPane.pane());
      case WorkbookLayoutCommand.SetSheetZoom setSheetZoom ->
          workbook.sheet(setSheetZoom.sheetName()).setZoom(setSheetZoom.zoomPercent());
      case WorkbookLayoutCommand.SetSheetPresentation setSheetPresentation ->
          workbook
              .sheet(setSheetPresentation.sheetName())
              .setPresentation(setSheetPresentation.presentation());
      case WorkbookLayoutCommand.SetPrintLayout setPrintLayout ->
          workbook.sheet(setPrintLayout.sheetName()).setPrintLayout(setPrintLayout.printLayout());
      case WorkbookLayoutCommand.ClearPrintLayout clearPrintLayout ->
          workbook.sheet(clearPrintLayout.sheetName()).clearPrintLayout();
      default -> throw new IllegalStateException("Unhandled sheet-structure command: " + command);
    }
  }

  static void applyCellValueCommand(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookCellCommand.SetCell setCell ->
          workbook.sheet(setCell.sheetName()).setCell(setCell.address(), setCell.value());
      case WorkbookCellCommand.SetRange setRange ->
          workbook.sheet(setRange.sheetName()).setRange(setRange.range(), setRange.rows());
      case WorkbookCellCommand.ClearRange clearRange ->
          workbook.sheet(clearRange.sheetName()).clearRange(clearRange.range());
      case WorkbookCellCommand.SetArrayFormula setArrayFormula ->
          workbook
              .sheet(setArrayFormula.sheetName())
              .setArrayFormula(setArrayFormula.range(), setArrayFormula.formula());
      case WorkbookCellCommand.ClearArrayFormula clearArrayFormula ->
          workbook
              .sheet(clearArrayFormula.sheetName())
              .clearArrayFormula(clearArrayFormula.address());
      case WorkbookCellCommand.AppendRow appendRow ->
          workbook
              .sheet(appendRow.sheetName())
              .appendRow(appendRow.values().toArray(ExcelCellValue[]::new));
      case WorkbookLayoutCommand.AutoSizeColumns autoSizeColumns ->
          workbook.sheet(autoSizeColumns.sheetName()).autoSizeColumns();
      default -> throw new IllegalStateException("Unhandled cell-value command: " + command);
    }
  }

  static void applyWorkbookMetadataCommand(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookAnnotationCommand.SetHyperlink setHyperlink ->
          workbook
              .sheet(setHyperlink.sheetName())
              .setHyperlink(setHyperlink.address(), setHyperlink.target());
      case WorkbookMetadataCommand.ImportCustomXmlMapping importCustomXmlMapping ->
          workbook.importCustomXmlMapping(importCustomXmlMapping.mapping());
      case WorkbookAnnotationCommand.ClearHyperlink clearHyperlink ->
          workbook.sheet(clearHyperlink.sheetName()).clearHyperlink(clearHyperlink.address());
      case WorkbookAnnotationCommand.SetComment setComment ->
          workbook
              .sheet(setComment.sheetName())
              .setComment(setComment.address(), setComment.comment());
      case WorkbookAnnotationCommand.ClearComment clearComment ->
          workbook.sheet(clearComment.sheetName()).clearComment(clearComment.address());
      case WorkbookDrawingCommand.SetPicture setPicture ->
          workbook.sheet(setPicture.sheetName()).setPicture(setPicture.picture());
      case WorkbookDrawingCommand.SetSignatureLine setSignatureLine ->
          workbook
              .sheet(setSignatureLine.sheetName())
              .setSignatureLine(setSignatureLine.signatureLine());
      case WorkbookDrawingCommand.SetChart setChart ->
          workbook.sheet(setChart.sheetName()).setChart(setChart.chart());
      case WorkbookDrawingCommand.SetShape setShape ->
          workbook.sheet(setShape.sheetName()).setShape(setShape.shape());
      case WorkbookDrawingCommand.SetEmbeddedObject setEmbeddedObject ->
          workbook
              .sheet(setEmbeddedObject.sheetName())
              .setEmbeddedObject(setEmbeddedObject.embeddedObject());
      case WorkbookDrawingCommand.SetDrawingObjectAnchor setDrawingObjectAnchor ->
          workbook
              .sheet(setDrawingObjectAnchor.sheetName())
              .setDrawingObjectAnchor(
                  setDrawingObjectAnchor.objectName(), setDrawingObjectAnchor.anchor());
      case WorkbookDrawingCommand.DeleteDrawingObject deleteDrawingObject ->
          workbook
              .sheet(deleteDrawingObject.sheetName())
              .deleteDrawingObject(deleteDrawingObject.objectName());
      case WorkbookFormattingCommand.ApplyStyle applyStyle ->
          workbook.sheet(applyStyle.sheetName()).applyStyle(applyStyle.range(), applyStyle.style());
      case WorkbookFormattingCommand.SetDataValidation setDataValidation ->
          workbook
              .sheet(setDataValidation.sheetName())
              .setDataValidation(setDataValidation.range(), setDataValidation.validation());
      case WorkbookFormattingCommand.ClearDataValidations clearDataValidations ->
          workbook
              .sheet(clearDataValidations.sheetName())
              .clearDataValidations(clearDataValidations.selection());
      case WorkbookFormattingCommand.SetConditionalFormatting setConditionalFormatting ->
          workbook
              .sheet(setConditionalFormatting.sheetName())
              .setConditionalFormatting(setConditionalFormatting.block());
      case WorkbookFormattingCommand.ClearConditionalFormatting clearConditionalFormatting ->
          workbook
              .sheet(clearConditionalFormatting.sheetName())
              .clearConditionalFormatting(clearConditionalFormatting.selection());
      case WorkbookTabularCommand.SetAutofilter setAutofilter ->
          workbook
              .sheet(setAutofilter.sheetName())
              .setAutofilter(
                  setAutofilter.range(), setAutofilter.criteria(), setAutofilter.sortState());
      case WorkbookTabularCommand.ClearAutofilter clearAutofilter ->
          workbook.sheet(clearAutofilter.sheetName()).clearAutofilter();
      case WorkbookTabularCommand.SetTable setTable -> workbook.setTable(setTable.definition());
      case WorkbookTabularCommand.SetPivotTable setPivotTable ->
          workbook.setPivotTable(setPivotTable.definition());
      case WorkbookTabularCommand.DeleteTable deleteTable ->
          workbook.deleteTable(deleteTable.name(), deleteTable.sheetName());
      case WorkbookTabularCommand.DeletePivotTable deletePivotTable ->
          workbook.deletePivotTable(deletePivotTable.name(), deletePivotTable.sheetName());
      case WorkbookMetadataCommand.SetNamedRange setNamedRange ->
          workbook.setNamedRange(setNamedRange.definition());
      case WorkbookMetadataCommand.DeleteNamedRange deleteNamedRange ->
          workbook.deleteNamedRange(deleteNamedRange.name(), deleteNamedRange.scope());
      default -> throw new IllegalStateException("Unhandled workbook-metadata command: " + command);
    }
  }
}
