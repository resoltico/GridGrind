package dev.erst.gridgrind.excel;

import java.util.Arrays;
import java.util.Objects;

/** Applies validated workbook commands to a workbook instance. */
final class WorkbookCommandExecutor {
  /** Applies one or more commands in order. */
  ExcelWorkbook apply(ExcelWorkbook workbook, WorkbookCommand... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return apply(workbook, Arrays.asList(commands));
  }

  /** Applies commands from any iterable source in order. */
  ExcelWorkbook apply(ExcelWorkbook workbook, Iterable<WorkbookCommand> commands) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    for (WorkbookCommand command : commands) {
      Objects.requireNonNull(command, "command must not be null");
      applyOne(workbook, command);
    }

    return workbook;
  }

  private void applyOne(ExcelWorkbook workbook, WorkbookCommand command) {
    switch (command) {
      case WorkbookSheetCommand sheetCommand -> applyWorkbookScopeCommand(workbook, sheetCommand);
      case WorkbookStructureCommand structureCommand ->
          applySheetStructureCommand(workbook, structureCommand);
      case WorkbookLayoutCommand layoutCommand -> applySheetLayoutCommand(workbook, layoutCommand);
      case WorkbookCellCommand cellCommand -> applyCellValueCommand(workbook, cellCommand);
      case WorkbookAnnotationCommand annotationCommand ->
          applyAnnotationCommand(workbook, annotationCommand);
      case WorkbookMetadataCommand metadataCommand ->
          applyMetadataCommand(workbook, metadataCommand);
      case WorkbookDrawingCommand drawingCommand -> applyDrawingCommand(workbook, drawingCommand);
      case WorkbookFormattingCommand formattingCommand ->
          applyFormattingCommand(workbook, formattingCommand);
      case WorkbookTabularCommand tabularCommand -> applyTabularCommand(workbook, tabularCommand);
    }
    workbook.markPackageMutated();
    workbook.invalidateFormulaRuntime();
  }

  static void applyWorkbookScopeCommand(ExcelWorkbook workbook, WorkbookSheetCommand command) {
    switch (command) {
      case WorkbookSheetCommand.CreateSheet createSheet ->
          workbook.getOrCreateSheet(createSheet.sheetName());
      case WorkbookSheetCommand.RenameSheet renameSheet ->
          workbook.sheets().renameSheet(renameSheet.sheetName(), renameSheet.newSheetName());
      case WorkbookSheetCommand.DeleteSheet deleteSheet ->
          workbook.sheets().deleteSheet(deleteSheet.sheetName());
      case WorkbookSheetCommand.MoveSheet moveSheet ->
          workbook.sheets().moveSheet(moveSheet.sheetName(), moveSheet.targetIndex());
      case WorkbookSheetCommand.CopySheet copySheet ->
          workbook
              .sheets()
              .copySheet(
                  copySheet.sourceSheetName(), copySheet.newSheetName(), copySheet.position());
      case WorkbookSheetCommand.SetActiveSheet setActiveSheet ->
          workbook.sheets().setActiveSheet(setActiveSheet.sheetName());
      case WorkbookSheetCommand.SetSelectedSheets setSelectedSheets ->
          workbook.sheets().setSelectedSheets(setSelectedSheets.sheetNames());
      case WorkbookSheetCommand.SetSheetVisibility setSheetVisibility ->
          workbook
              .sheets()
              .setSheetVisibility(setSheetVisibility.sheetName(), setSheetVisibility.visibility());
      case WorkbookSheetCommand.SetSheetProtection setSheetProtection ->
          workbook
              .sheets()
              .setSheetProtection(
                  setSheetProtection.sheetName(),
                  setSheetProtection.protection(),
                  setSheetProtection.password());
      case WorkbookSheetCommand.ClearSheetProtection clearSheetProtection ->
          workbook.sheets().clearSheetProtection(clearSheetProtection.sheetName());
      case WorkbookSheetCommand.SetWorkbookProtection setWorkbookProtection ->
          workbook.protection().setWorkbookProtection(setWorkbookProtection.protection());
      case WorkbookSheetCommand.ClearWorkbookProtection _ ->
          workbook.protection().clearWorkbookProtection();
    }
  }

  static void applySheetStructureCommand(ExcelWorkbook workbook, WorkbookStructureCommand command) {
    switch (command) {
      case WorkbookStructureCommand.MergeCells mergeCells ->
          workbook.sheet(mergeCells.sheetName()).layout().mergeCells(mergeCells.range());
      case WorkbookStructureCommand.UnmergeCells unmergeCells ->
          workbook.sheet(unmergeCells.sheetName()).layout().unmergeCells(unmergeCells.range());
      case WorkbookStructureCommand.SetColumnWidth setColumnWidth ->
          workbook
              .sheet(setColumnWidth.sheetName())
              .columns()
              .setWidth(
                  setColumnWidth.firstColumnIndex(),
                  setColumnWidth.lastColumnIndex(),
                  setColumnWidth.widthCharacters());
      case WorkbookStructureCommand.SetRowHeight setRowHeight ->
          workbook
              .sheet(setRowHeight.sheetName())
              .rows()
              .setHeight(
                  setRowHeight.firstRowIndex(),
                  setRowHeight.lastRowIndex(),
                  setRowHeight.heightPoints());
      case WorkbookStructureCommand.InsertRows insertRows ->
          workbook
              .sheet(insertRows.sheetName())
              .rows()
              .insertRows(insertRows.rowIndex(), insertRows.rowCount());
      case WorkbookStructureCommand.DeleteRows deleteRows ->
          workbook.sheet(deleteRows.sheetName()).rows().deleteRows(deleteRows.rows());
      case WorkbookStructureCommand.ShiftRows shiftRows ->
          workbook
              .sheet(shiftRows.sheetName())
              .rows()
              .shiftRows(shiftRows.rows(), shiftRows.delta());
      case WorkbookStructureCommand.InsertColumns insertColumns ->
          workbook
              .sheet(insertColumns.sheetName())
              .columns()
              .insert(insertColumns.columnIndex(), insertColumns.columnCount());
      case WorkbookStructureCommand.DeleteColumns deleteColumns ->
          workbook.sheet(deleteColumns.sheetName()).columns().delete(deleteColumns.columns());
      case WorkbookStructureCommand.ShiftColumns shiftColumns ->
          workbook
              .sheet(shiftColumns.sheetName())
              .columns()
              .shift(shiftColumns.columns(), shiftColumns.delta());
      case WorkbookStructureCommand.SetRowVisibility setRowVisibility ->
          workbook
              .sheet(setRowVisibility.sheetName())
              .rows()
              .setVisibility(setRowVisibility.rows(), setRowVisibility.hidden());
      case WorkbookStructureCommand.SetColumnVisibility setColumnVisibility ->
          workbook
              .sheet(setColumnVisibility.sheetName())
              .columns()
              .setVisibility(setColumnVisibility.columns(), setColumnVisibility.hidden());
      case WorkbookStructureCommand.GroupRows groupRows ->
          workbook
              .sheet(groupRows.sheetName())
              .rows()
              .group(groupRows.rows(), groupRows.collapsed());
      case WorkbookStructureCommand.UngroupRows ungroupRows ->
          workbook.sheet(ungroupRows.sheetName()).rows().ungroup(ungroupRows.rows());
      case WorkbookStructureCommand.GroupColumns groupColumns ->
          workbook
              .sheet(groupColumns.sheetName())
              .columns()
              .group(groupColumns.columns(), groupColumns.collapsed());
      case WorkbookStructureCommand.UngroupColumns ungroupColumns ->
          workbook.sheet(ungroupColumns.sheetName()).columns().ungroup(ungroupColumns.columns());
    }
  }

  static void applySheetLayoutCommand(ExcelWorkbook workbook, WorkbookLayoutCommand command) {
    switch (command) {
      case WorkbookLayoutCommand.SetSheetPane setSheetPane ->
          workbook.sheet(setSheetPane.sheetName()).layout().setPane(setSheetPane.pane());
      case WorkbookLayoutCommand.SetSheetZoom setSheetZoom ->
          workbook.sheet(setSheetZoom.sheetName()).layout().setZoom(setSheetZoom.zoomPercent());
      case WorkbookLayoutCommand.SetSheetPresentation setSheetPresentation ->
          workbook
              .sheet(setSheetPresentation.sheetName())
              .layout()
              .setPresentation(setSheetPresentation.presentation());
      case WorkbookLayoutCommand.SetPrintLayout setPrintLayout ->
          workbook
              .sheet(setPrintLayout.sheetName())
              .layout()
              .setPrintLayout(setPrintLayout.printLayout());
      case WorkbookLayoutCommand.ClearPrintLayout clearPrintLayout ->
          workbook.sheet(clearPrintLayout.sheetName()).layout().clearPrintLayout();
      case WorkbookLayoutCommand.AutoSizeColumns autoSizeColumns ->
          applyAutoSizeColumnsCommand(workbook, autoSizeColumns);
    }
  }

  static void applyCellValueCommand(ExcelWorkbook workbook, WorkbookCellCommand command) {
    switch (command) {
      case WorkbookCellCommand.SetCell setCell ->
          workbook.sheet(setCell.sheetName()).cells().setCell(setCell.address(), setCell.value());
      case WorkbookCellCommand.SetRange setRange ->
          workbook.sheet(setRange.sheetName()).cells().setRange(setRange.range(), setRange.rows());
      case WorkbookCellCommand.ClearRange clearRange ->
          workbook.sheet(clearRange.sheetName()).cells().clearRange(clearRange.range());
      case WorkbookCellCommand.SetArrayFormula setArrayFormula ->
          workbook
              .sheet(setArrayFormula.sheetName())
              .cells()
              .setArrayFormula(setArrayFormula.range(), setArrayFormula.formula());
      case WorkbookCellCommand.ClearArrayFormula clearArrayFormula ->
          workbook
              .sheet(clearArrayFormula.sheetName())
              .cells()
              .clearArrayFormula(clearArrayFormula.address());
      case WorkbookCellCommand.AppendRow appendRow ->
          workbook
              .sheet(appendRow.sheetName())
              .cells()
              .appendRow(appendRow.values().toArray(ExcelCellValue[]::new));
    }
  }

  private static void applyAutoSizeColumnsCommand(
      ExcelWorkbook workbook, WorkbookLayoutCommand.AutoSizeColumns command) {
    workbook.sheet(command.sheetName()).columns().autoSize();
  }

  static void applyAnnotationCommand(ExcelWorkbook workbook, WorkbookAnnotationCommand command) {
    switch (command) {
      case WorkbookAnnotationCommand.SetHyperlink setHyperlink ->
          workbook
              .sheet(setHyperlink.sheetName())
              .annotations()
              .setHyperlink(setHyperlink.address(), setHyperlink.target());
      case WorkbookAnnotationCommand.ClearHyperlink clearHyperlink ->
          workbook
              .sheet(clearHyperlink.sheetName())
              .annotations()
              .clearHyperlink(clearHyperlink.address());
      case WorkbookAnnotationCommand.SetComment setComment ->
          workbook
              .sheet(setComment.sheetName())
              .annotations()
              .setComment(setComment.address(), setComment.comment());
      case WorkbookAnnotationCommand.ClearComment clearComment ->
          workbook
              .sheet(clearComment.sheetName())
              .annotations()
              .clearComment(clearComment.address());
    }
  }

  static void applyMetadataCommand(ExcelWorkbook workbook, WorkbookMetadataCommand command) {
    switch (command) {
      case WorkbookMetadataCommand.ImportCustomXmlMapping importCustomXmlMapping ->
          workbook.customXml().importCustomXmlMapping(importCustomXmlMapping.mapping());
      case WorkbookMetadataCommand.SetNamedRange setNamedRange ->
          workbook.names().setNamedRange(setNamedRange.definition());
      case WorkbookMetadataCommand.DeleteNamedRange deleteNamedRange ->
          workbook.names().deleteNamedRange(deleteNamedRange.name(), deleteNamedRange.scope());
    }
  }

  static void applyDrawingCommand(ExcelWorkbook workbook, WorkbookDrawingCommand command) {
    switch (command) {
      case WorkbookDrawingCommand.SetPicture setPicture ->
          workbook.sheet(setPicture.sheetName()).drawings().setPicture(setPicture.picture());
      case WorkbookDrawingCommand.SetSignatureLine setSignatureLine ->
          workbook
              .sheet(setSignatureLine.sheetName())
              .drawings()
              .setSignatureLine(setSignatureLine.signatureLine());
      case WorkbookDrawingCommand.SetChart setChart ->
          workbook.sheet(setChart.sheetName()).drawings().setChart(setChart.chart());
      case WorkbookDrawingCommand.SetShape setShape ->
          workbook.sheet(setShape.sheetName()).drawings().setShape(setShape.shape());
      case WorkbookDrawingCommand.SetEmbeddedObject setEmbeddedObject ->
          workbook
              .sheet(setEmbeddedObject.sheetName())
              .drawings()
              .setEmbeddedObject(setEmbeddedObject.embeddedObject());
      case WorkbookDrawingCommand.SetDrawingObjectAnchor setDrawingObjectAnchor ->
          workbook
              .sheet(setDrawingObjectAnchor.sheetName())
              .drawings()
              .setDrawingObjectAnchor(
                  setDrawingObjectAnchor.objectName(), setDrawingObjectAnchor.anchor());
      case WorkbookDrawingCommand.DeleteDrawingObject deleteDrawingObject ->
          workbook
              .sheet(deleteDrawingObject.sheetName())
              .drawings()
              .deleteDrawingObject(deleteDrawingObject.objectName());
    }
  }

  static void applyFormattingCommand(ExcelWorkbook workbook, WorkbookFormattingCommand command) {
    switch (command) {
      case WorkbookFormattingCommand.ApplyStyle applyStyle ->
          workbook
              .sheet(applyStyle.sheetName())
              .cells()
              .applyStyle(applyStyle.range(), applyStyle.style());
      case WorkbookFormattingCommand.SetDataValidation setDataValidation ->
          workbook
              .sheet(setDataValidation.sheetName())
              .metadata()
              .setDataValidation(setDataValidation.range(), setDataValidation.validation());
      case WorkbookFormattingCommand.ClearDataValidations clearDataValidations ->
          workbook
              .sheet(clearDataValidations.sheetName())
              .metadata()
              .clearDataValidations(clearDataValidations.selection());
      case WorkbookFormattingCommand.SetConditionalFormatting setConditionalFormatting ->
          workbook
              .sheet(setConditionalFormatting.sheetName())
              .metadata()
              .setConditionalFormatting(setConditionalFormatting.block());
      case WorkbookFormattingCommand.ClearConditionalFormatting clearConditionalFormatting ->
          workbook
              .sheet(clearConditionalFormatting.sheetName())
              .metadata()
              .clearConditionalFormatting(clearConditionalFormatting.selection());
    }
  }

  static void applyTabularCommand(ExcelWorkbook workbook, WorkbookTabularCommand command) {
    switch (command) {
      case WorkbookTabularCommand.SetAutofilter setAutofilter ->
          workbook
              .sheet(setAutofilter.sheetName())
              .metadata()
              .setAutofilter(
                  setAutofilter.range(), setAutofilter.criteria(), setAutofilter.sortState());
      case WorkbookTabularCommand.ClearAutofilter clearAutofilter ->
          workbook.sheet(clearAutofilter.sheetName()).metadata().clearAutofilter();
      case WorkbookTabularCommand.SetTable setTable ->
          workbook.tables().setTable(setTable.definition());
      case WorkbookTabularCommand.SetPivotTable setPivotTable ->
          workbook.pivots().setPivotTable(setPivotTable.definition());
      case WorkbookTabularCommand.DeleteTable deleteTable ->
          workbook.tables().deleteTable(deleteTable.name(), deleteTable.sheetName());
      case WorkbookTabularCommand.DeletePivotTable deletePivotTable ->
          workbook.pivots().deletePivotTable(deletePivotTable.name(), deletePivotTable.sheetName());
    }
  }
}
