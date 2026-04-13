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
                  + " FORCE_FORMULA_RECALC_ON_OPEN only; unsupported operation type: "
                  + commandType(command));
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

  private static String commandType(WorkbookCommand command) {
    return switch (command) {
      case WorkbookCommand.CreateSheet _ -> "ENSURE_SHEET";
      case WorkbookCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookCommand.ForceFormulaRecalculationOnOpen _ -> "FORCE_FORMULA_RECALC_ON_OPEN";
      case WorkbookCommand.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookCommand.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookCommand.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookCommand.CopySheet _ -> "COPY_SHEET";
      case WorkbookCommand.SetActiveSheet _ -> "SET_ACTIVE_SHEET";
      case WorkbookCommand.SetSelectedSheets _ -> "SET_SELECTED_SHEETS";
      case WorkbookCommand.SetSheetVisibility _ -> "SET_SHEET_VISIBILITY";
      case WorkbookCommand.SetSheetProtection _ -> "SET_SHEET_PROTECTION";
      case WorkbookCommand.ClearSheetProtection _ -> "CLEAR_SHEET_PROTECTION";
      case WorkbookCommand.SetWorkbookProtection _ -> "SET_WORKBOOK_PROTECTION";
      case WorkbookCommand.ClearWorkbookProtection _ -> "CLEAR_WORKBOOK_PROTECTION";
      case WorkbookCommand.MergeCells _ -> "MERGE_CELLS";
      case WorkbookCommand.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookCommand.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookCommand.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookCommand.InsertRows _ -> "INSERT_ROWS";
      case WorkbookCommand.DeleteRows _ -> "DELETE_ROWS";
      case WorkbookCommand.ShiftRows _ -> "SHIFT_ROWS";
      case WorkbookCommand.InsertColumns _ -> "INSERT_COLUMNS";
      case WorkbookCommand.DeleteColumns _ -> "DELETE_COLUMNS";
      case WorkbookCommand.ShiftColumns _ -> "SHIFT_COLUMNS";
      case WorkbookCommand.SetRowVisibility _ -> "SET_ROW_VISIBILITY";
      case WorkbookCommand.SetColumnVisibility _ -> "SET_COLUMN_VISIBILITY";
      case WorkbookCommand.GroupRows _ -> "GROUP_ROWS";
      case WorkbookCommand.UngroupRows _ -> "UNGROUP_ROWS";
      case WorkbookCommand.GroupColumns _ -> "GROUP_COLUMNS";
      case WorkbookCommand.UngroupColumns _ -> "UNGROUP_COLUMNS";
      case WorkbookCommand.SetSheetPane _ -> "SET_SHEET_PANE";
      case WorkbookCommand.SetSheetZoom _ -> "SET_SHEET_ZOOM";
      case WorkbookCommand.SetSheetPresentation _ -> "SET_SHEET_PRESENTATION";
      case WorkbookCommand.SetPrintLayout _ -> "SET_PRINT_LAYOUT";
      case WorkbookCommand.ClearPrintLayout _ -> "CLEAR_PRINT_LAYOUT";
      case WorkbookCommand.SetCell _ -> "SET_CELL";
      case WorkbookCommand.SetRange _ -> "SET_RANGE";
      case WorkbookCommand.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookCommand.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookCommand.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookCommand.SetComment _ -> "SET_COMMENT";
      case WorkbookCommand.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookCommand.SetPicture _ -> "SET_PICTURE";
      case WorkbookCommand.SetChart _ -> "SET_CHART";
      case WorkbookCommand.SetPivotTable _ -> "SET_PIVOT_TABLE";
      case WorkbookCommand.SetShape _ -> "SET_SHAPE";
      case WorkbookCommand.SetEmbeddedObject _ -> "SET_EMBEDDED_OBJECT";
      case WorkbookCommand.SetDrawingObjectAnchor _ -> "SET_DRAWING_OBJECT_ANCHOR";
      case WorkbookCommand.DeleteDrawingObject _ -> "DELETE_DRAWING_OBJECT";
      case WorkbookCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookCommand.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookCommand.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookCommand.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookCommand.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookCommand.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookCommand.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookCommand.SetTable _ -> "SET_TABLE";
      case WorkbookCommand.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookCommand.DeletePivotTable _ -> "DELETE_PIVOT_TABLE";
      case WorkbookCommand.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookCommand.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case WorkbookCommand.EvaluateAllFormulas _ -> "EVALUATE_FORMULAS";
      case WorkbookCommand.EvaluateFormulaCells _ -> "EVALUATE_FORMULA_CELLS";
      case WorkbookCommand.ClearFormulaCaches _ -> "CLEAR_FORMULA_CACHES";
    };
  }
}
