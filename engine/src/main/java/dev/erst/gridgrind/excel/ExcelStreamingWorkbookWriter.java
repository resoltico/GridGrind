package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/** Stateful low-memory append-oriented workbook authoring backed by POI's SXSSF writer. */
public final class ExcelStreamingWorkbookWriter implements AutoCloseable {
  private final SXSSFWorkbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final Map<String, SXSSFSheet> sheets;
  private final Map<String, Integer> nextRowIndexes;

  /** Creates one empty streaming workbook session. */
  public ExcelStreamingWorkbookWriter() {
    this.workbook = new SXSSFWorkbook(100);
    this.workbook.setCompressTempFiles(true);
    this.styleRegistry = new WorkbookStyleRegistry(workbook.getXSSFWorkbook());
    this.sheets = new HashMap<>();
    this.nextRowIndexes = new HashMap<>();
  }

  /** Applies one supported streaming command to the current workbook session. */
  public void apply(WorkbookCommand command) {
    Objects.requireNonNull(command, "command must not be null");
    switch (command) {
      case WorkbookCommand.CreateSheet createSheet ->
          sheets.computeIfAbsent(
              createSheet.sheetName(),
              sheetName -> {
                nextRowIndexes.put(sheetName, 0);
                return workbook.createSheet(sheetName);
              });
      case WorkbookCommand.AppendRow appendRow ->
          appendRow(requiredSheet(appendRow.sheetName()), appendRow);
      case WorkbookCommand.ForceFormulaRecalculationOnOpen _ ->
          workbook.setForceFormulaRecalculation(true);
      default ->
          throw new IllegalArgumentException(
              "executionMode.writeMode=STREAMING_WRITE supports ENSURE_SHEET, APPEND_ROW, and"
                  + " FORCE_FORMULA_RECALCULATION_ON_OPEN only; unsupported operation type: "
                  + command.commandType());
    }
  }

  /** Materializes the current streaming workbook session to one `.xlsx` path. */
  public void save(Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Path absolutePath = workbookPath.toAbsolutePath().normalize();
    Files.createDirectories(absolutePath.getParent());
    try (OutputStream outputStream = Files.newOutputStream(absolutePath)) {
      workbook.write(outputStream);
    }
  }

  @Override
  public void close() throws IOException {
    try {
      workbook.close();
    } finally {
      workbook.dispose();
    }
  }

  private SXSSFSheet requiredSheet(String sheetName) {
    SXSSFSheet sheet = sheets.get(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  private void appendRow(SXSSFSheet sheet, WorkbookCommand.AppendRow appendRow) {
    int rowIndex =
        nextRowIndexes.compute(appendRow.sheetName(), (_, nextRowIndex) -> nextRowIndex + 1) - 1;
    SXSSFRow row = sheet.createRow(rowIndex);
    for (int columnIndex = 0; columnIndex < appendRow.values().size(); columnIndex++) {
      writeCellValue(row.createCell(columnIndex), appendRow.values().get(columnIndex));
    }
  }

  private void writeCellValue(SXSSFCell cell, ExcelCellValue value) {
    switch (value) {
      case ExcelCellValue.BlankValue _ -> cell.setBlank();
      case ExcelCellValue.TextValue textValue -> cell.setCellValue(textValue.value());
      case ExcelCellValue.RichTextValue richTextValue ->
          cell.setCellValue(
              ExcelRichTextSupport.toPoiRichText(
                  workbook.getXSSFWorkbook(), richTextValue.value()));
      case ExcelCellValue.NumberValue numberValue -> cell.setCellValue(numberValue.value());
      case ExcelCellValue.BooleanValue booleanValue -> cell.setCellValue(booleanValue.value());
      case ExcelCellValue.DateValue dateValue -> {
        cell.setCellValue(dateValue.value());
        cell.setCellStyle(styleRegistry.localDateStyle(cell));
      }
      case ExcelCellValue.DateTimeValue dateTimeValue -> {
        cell.setCellValue(dateTimeValue.value());
        cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      }
      case ExcelCellValue.FormulaValue formulaValue -> {
        cell.setCellFormula(formulaValue.expression());
        clearPersistedFormulaCache(cell);
      }
    }
  }

  private static void clearPersistedFormulaCache(Cell cell) {
    cell.setCellType(CellType.FORMULA);
  }
}
