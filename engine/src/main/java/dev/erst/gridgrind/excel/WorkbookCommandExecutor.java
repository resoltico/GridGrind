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
      case WorkbookCommand.SetCell setCell ->
          workbook
              .getOrCreateSheet(setCell.sheetName())
              .setCell(setCell.address(), setCell.value());
      case WorkbookCommand.SetRange setRange ->
          workbook
              .getOrCreateSheet(setRange.sheetName())
              .setRange(setRange.range(), setRange.rows());
      case WorkbookCommand.ClearRange clearRange ->
          workbook.sheet(clearRange.sheetName()).clearRange(clearRange.range());
      case WorkbookCommand.ApplyStyle applyStyle ->
          workbook
              .getOrCreateSheet(applyStyle.sheetName())
              .applyStyle(applyStyle.range(), applyStyle.style());
      case WorkbookCommand.AppendRow appendRow ->
          workbook
              .getOrCreateSheet(appendRow.sheetName())
              .appendRow(appendRow.values().toArray(ExcelCellValue[]::new));
      case WorkbookCommand.AutoSizeColumns autoSizeColumns ->
          workbook
              .sheet(autoSizeColumns.sheetName())
              .autoSizeColumns(autoSizeColumns.columnNames().toArray(String[]::new));
      case WorkbookCommand.EvaluateAllFormulas _ -> workbook.evaluateAllFormulas();
      case WorkbookCommand.ForceFormulaRecalculationOnOpen _ ->
          workbook.forceFormulaRecalculationOnOpen();
    }
  }
}
