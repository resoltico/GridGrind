package dev.erst.gridgrind.excel.stream;

import dev.erst.gridgrind.excel.ExcelCellErrorLiteralSupport;
import dev.erst.gridgrind.excel.ExcelCellTextLimits;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelDeterministicWorkbookArtifactSupport;
import dev.erst.gridgrind.excel.ExcelFormulaWriteSupport;
import dev.erst.gridgrind.excel.ExcelRichTextSupport;
import dev.erst.gridgrind.excel.ExcelWorkbookDocumentMetadataSupport;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookCellCommand;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookSheetCommand;
import dev.erst.gridgrind.excel.WorkbookStyleRegistry;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageFileSupport;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Stateful low-memory append-oriented workbook authoring backed by POI's SXSSF writer. */
public final class ExcelStreamingWorkbookWriter implements AutoCloseable {
  private final SXSSFWorkbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final Map<String, SXSSFSheet> sheets;
  private final Map<String, Integer> nextRowIndexes;

  /** Creates one empty streaming workbook session. */
  public ExcelStreamingWorkbookWriter() {
    this.workbook = new SXSSFWorkbook(new XSSFWorkbook(), 100, true, true);
    this.styleRegistry = new WorkbookStyleRegistry(workbook.getXSSFWorkbook());
    this.sheets = new HashMap<>();
    this.nextRowIndexes = new HashMap<>();
  }

  /** Applies one supported streaming command to the current workbook session. */
  public void apply(WorkbookCommand command) {
    Objects.requireNonNull(command, "command must not be null");
    switch (command) {
      case WorkbookSheetCommand.CreateSheet createSheet ->
          sheets.computeIfAbsent(
              createSheet.sheetName(),
              sheetName -> {
                nextRowIndexes.put(sheetName, 0);
                return workbook.createSheet(sheetName);
              });
      case WorkbookCellCommand.AppendRow appendRow ->
          appendRow(requiredSheet(appendRow.sheetName()), appendRow);
      default ->
          throw new IllegalArgumentException(
              "execution.mode.type=STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW"
                  + " only; unsupported mutation action type: "
                  + command.commandType());
    }
  }

  /** Marks the streamed workbook to recalculate formulas on the next Excel-compatible open. */
  public void markRecalculateOnOpen() {
    workbook.setForceFormulaRecalculation(true);
  }

  /** Materializes the current streaming workbook session to one `.xlsx` path. */
  public void save(Path workbookPath, WorkbookArtifactWriteDisposition writeDisposition)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    Path absolutePath = workbookPath.toAbsolutePath().normalize();
    Files.createDirectories(absolutePath.getParent());
    ExcelWorkbookDocumentMetadataSupport.normalizeForSave(workbook.getXSSFWorkbook());
    try (OutputStream outputStream =
        ExcelOoxmlPackageFileSupport.newWorkbookOutputStream(absolutePath, writeDisposition)) {
      workbook.write(outputStream);
    }
    ExcelDeterministicWorkbookArtifactSupport.normalizeWorkbookPackage(absolutePath);
  }

  @Override
  public void close() throws IOException {
    workbook.close();
  }

  private SXSSFSheet requiredSheet(String sheetName) {
    SXSSFSheet sheet = sheets.get(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  private void appendRow(SXSSFSheet sheet, WorkbookCellCommand.AppendRow appendRow) {
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
      case ExcelCellValue.TextValue textValue -> {
        ExcelCellTextLimits.requireSupportedLength(textValue.value(), "textValue.value"); // LIM-010
        cell.setCellValue(textValue.value());
      }
      case ExcelCellValue.RichTextValue richTextValue ->
          cell.setCellValue(
              ExcelRichTextSupport.toPoiRichText(
                  workbook.getXSSFWorkbook(), richTextValue.value()));
      case ExcelCellValue.NumberValue numberValue -> cell.setCellValue(numberValue.value());
      case ExcelCellValue.BooleanValue booleanValue -> cell.setCellValue(booleanValue.value());
      case ExcelCellValue.ErrorValue errorValue ->
          cell.setCellErrorValue(
              ExcelCellErrorLiteralSupport.toPoiStoredFormulaError(errorValue.value()).getCode());
      case ExcelCellValue.DateValue dateValue -> {
        cell.setCellValue(dateValue.value());
        cell.setCellStyle(styleRegistry.localDateStyle(cell));
      }
      case ExcelCellValue.DateTimeValue dateTimeValue -> {
        cell.setCellValue(dateTimeValue.value());
        cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      }
      case ExcelCellValue.FormulaValue formulaValue ->
          ExcelFormulaWriteSupport.setAuthoredFormula(cell, formulaValue.expression());
      case ExcelCellValue.RawFormulaValue rawFormulaValue ->
          ExcelFormulaWriteSupport.setOpaqueFormula(cell, rawFormulaValue.expression());
    }
  }
}
