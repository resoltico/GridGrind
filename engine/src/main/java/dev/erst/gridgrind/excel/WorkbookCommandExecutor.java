package dev.erst.gridgrind.excel;

import java.util.Arrays;
import java.util.Objects;

/** Applies validated workbook commands to a workbook instance. */
public final class WorkbookCommandExecutor {
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
              setSheetProtection.sheetName(), setSheetProtection.protection());
      case WorkbookCommand.ClearSheetProtection clearSheetProtection ->
          workbook.clearSheetProtection(clearSheetProtection.sheetName());
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
      case WorkbookCommand.SetSheetPane setSheetPane ->
          workbook.sheet(setSheetPane.sheetName()).setPane(setSheetPane.pane());
      case WorkbookCommand.SetSheetZoom setSheetZoom ->
          workbook.sheet(setSheetZoom.sheetName()).setZoom(setSheetZoom.zoomPercent());
      case WorkbookCommand.SetPrintLayout setPrintLayout ->
          workbook.sheet(setPrintLayout.sheetName()).setPrintLayout(setPrintLayout.printLayout());
      case WorkbookCommand.ClearPrintLayout clearPrintLayout ->
          workbook.sheet(clearPrintLayout.sheetName()).clearPrintLayout();
      case WorkbookCommand.SetCell setCell ->
          workbook.sheet(setCell.sheetName()).setCell(setCell.address(), setCell.value());
      case WorkbookCommand.SetRange setRange ->
          workbook.sheet(setRange.sheetName()).setRange(setRange.range(), setRange.rows());
      case WorkbookCommand.ClearRange clearRange ->
          workbook.sheet(clearRange.sheetName()).clearRange(clearRange.range());
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
          workbook.sheet(setAutofilter.sheetName()).setAutofilter(setAutofilter.range());
      case WorkbookCommand.ClearAutofilter clearAutofilter ->
          workbook.sheet(clearAutofilter.sheetName()).clearAutofilter();
      case WorkbookCommand.SetTable setTable -> workbook.setTable(setTable.definition());
      case WorkbookCommand.DeleteTable deleteTable ->
          workbook.deleteTable(deleteTable.name(), deleteTable.sheetName());
      case WorkbookCommand.SetNamedRange setNamedRange ->
          workbook.setNamedRange(setNamedRange.definition());
      case WorkbookCommand.DeleteNamedRange deleteNamedRange ->
          workbook.deleteNamedRange(deleteNamedRange.name(), deleteNamedRange.scope());
      case WorkbookCommand.AppendRow appendRow ->
          workbook
              .sheet(appendRow.sheetName())
              .appendRow(appendRow.values().toArray(ExcelCellValue[]::new));
      case WorkbookCommand.AutoSizeColumns autoSizeColumns ->
          workbook.sheet(autoSizeColumns.sheetName()).autoSizeColumns();
      case WorkbookCommand.EvaluateAllFormulas _ -> workbook.evaluateAllFormulas();
      case WorkbookCommand.ForceFormulaRecalculationOnOpen _ ->
          workbook.forceFormulaRecalculationOnOpen();
    }
  }
}
