package dev.erst.gridgrind.engine.layout;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelIndexDisplay;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.engine.comments.CommentAdapter;
import dev.erst.gridgrind.engine.names.NamedRangeAdapter;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.poi.SqrefRanges;
import dev.erst.gridgrind.engine.validation.DataValidationAdapter;
import java.util.ArrayList;
import java.util.BitSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Owns structural row/column edits, safety guards and inserted formatting propagation. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit
      .ID_DESTRUCTIVE_STRUCTURAL_EDITS_AGAINST_RANGE_BACKED_NAMED_RANGES
})
public final class StructuralEdits {
  private StructuralEdits() {}

  public static void insertRows(XSSFSheet sheet, int rowIndex, int rowCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertRowBounds(sheet, rowIndex, rowCount);
    int lastRowIndex = sheet.getLastRowNum();
    List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation> expectedValidations =
        dev.erst.gridgrind.engine.validation.DataValidationAdapter
            .expectedValidationsAfterInsertRows(sheet, rowIndex, rowCount);
    rejectAffectedRowStructuresForInsert(sheet, rowIndex);
    if (rowIndex <= lastRowIndex) {
      sheet.shiftRows(rowIndex, lastRowIndex, rowCount, true, false);
      copyAdjacentVisualFormattingIntoInsertedRows(sheet, rowIndex, rowCount, lastRowIndex);
    }
    dev.erst.gridgrind.engine.validation.DataValidationAdapter.replaceDataValidations(
        sheet, expectedValidations);
  }

  public static void deleteRows(XSSFSheet sheet, ExcelRowSpan rows) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    int lastRowIndex = sheet.getLastRowNum();
    if (lastRowIndex < 0) {
      throw new IllegalArgumentException("DELETE_ROWS requires at least one existing row");
    }
    if (rows.lastRowIndex() > lastRowIndex) {
      throw new IllegalArgumentException(
          "DELETE_ROWS rows must stay within existing row bounds: last existing row is "
              + ExcelIndexDisplay.rowValue(lastRowIndex)
              + "; requested "
              + ExcelIndexDisplay.describe("lastRowIndex", rows.lastRowIndex()));
    }
    rejectAffectedRowStructuresForDelete(sheet, rows);
    if (rows.lastRowIndex() < lastRowIndex) {
      sheet.shiftRows(rows.lastRowIndex() + 1, lastRowIndex, -rows.count(), true, false);
    }
    int clearStart = Math.max(rows.firstRowIndex(), lastRowIndex - rows.count() + 1);
    for (int rowIndex = clearStart; rowIndex <= lastRowIndex; rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      if (row != null) {
        sheet.removeRow(row);
      }
    }
  }

  public static void shiftRows(XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    requireShiftedRowBounds(rows, delta);
    rejectAffectedRowStructuresForShift(sheet, rows, delta);
    sheet.shiftRows(rows.firstRowIndex(), rows.lastRowIndex(), delta, true, false);
  }

  public static void setRowVisibility(XSSFSheet sheet, ExcelRowSpan rows, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    LayoutAdapter.ensureRowsExist(sheet, rows);
    for (int rowIndex = rows.firstRowIndex(); rowIndex <= rows.lastRowIndex(); rowIndex++) {
      sheet.getRow(rowIndex).setZeroHeight(hidden);
    }
  }

  public static void groupRows(XSSFSheet sheet, ExcelRowSpan rows, boolean collapsed) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    LayoutAdapter.ensureRowsExist(sheet, rows);
    sheet.groupRow(rows.firstRowIndex(), rows.lastRowIndex());
    if (collapsed) {
      LayoutAdapter.collapseRows(sheet, rows);
      return;
    }
    if (LayoutAdapter.isRowGroupCollapsed(sheet, rows)) {
      LayoutAdapter.expandRows(sheet, rows);
      return;
    }
    LayoutAdapter.clearExpandedGroupControlRow(sheet, rows);
  }

  public static void ungroupRows(XSSFSheet sheet, ExcelRowSpan rows) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    LayoutAdapter.ensureRowsExist(sheet, rows);
    if (LayoutAdapter.isRowGroupCollapsed(sheet, rows)) {
      LayoutAdapter.expandRows(sheet, rows);
    } else {
      LayoutAdapter.clearExpandedGroupControlRow(sheet, rows);
    }
    sheet.ungroupRow(rows.firstRowIndex(), rows.lastRowIndex());
  }

  public static List<RowLayoutReport> rowLayouts(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return LayoutAdapter.rowLayouts(sheet);
  }

  private static void requireInsertRowBounds(XSSFSheet sheet, int rowIndex, int rowCount) {
    int lastRowIndex = sheet.getLastRowNum();
    if (rowIndex > lastRowIndex + 1) {
      throw new IllegalArgumentException(
          "INSERT_ROWS "
              + ExcelIndexDisplay.describe("rowIndex", rowIndex)
              + " must be less than or equal to last existing row + 1: "
              + ExcelIndexDisplay.rowValue(lastRowIndex + 1));
    }
    if (rowIndex + rowCount - 1 > Limit.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_ROWS would exceed the maximum row index: destination last row would be "
              + ExcelIndexDisplay.rowValue(rowIndex + rowCount - 1)
              + "; maximum is "
              + ExcelIndexDisplay.rowValue(Limit.MAX_ROW_INDEX));
    }
  }

  private static void requireShiftedRowBounds(ExcelRowSpan rows, int delta) {
    if (rows.firstRowIndex() + delta < 0) {
      throw new IllegalArgumentException(
          "SHIFT_ROWS would move "
              + ExcelIndexDisplay.describe("firstRowIndex", rows.firstRowIndex())
              + " by delta "
              + delta
              + " before the first worksheet row ("
              + ExcelIndexDisplay.excelRow(0)
              + ")");
    }
    if (rows.lastRowIndex() + delta > Limit.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "SHIFT_ROWS would move "
              + ExcelIndexDisplay.describe("lastRowIndex", rows.lastRowIndex())
              + " by delta "
              + delta
              + " beyond the maximum row "
              + ExcelIndexDisplay.rowValue(Limit.MAX_ROW_INDEX));
    }
  }

  public static void insertColumns(XSSFSheet sheet, int columnIndex, int columnCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireInsertColumnBounds(sheet, columnIndex, columnCount);
    rejectFormulaBearingWorkbookForColumnEdit(sheet.getWorkbook(), "INSERT_COLUMNS");
    LayoutAdapter.normalizeColumnDefinitionContainer(sheet);
    int lastColumnIndex = lastColumnIndex(sheet);
    Map<Integer, CTCol> explicitColumns = LayoutAdapter.snapshotColumnDefinitions(sheet);
    List<CTDataValidation> expectedValidations =
        DataValidationAdapter.expectedValidationsAfterInsertColumns(
            sheet, columnIndex, columnCount);
    CommentAdapter commentRepairSupport = new CommentAdapter(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<CommentAdapter.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments =
          commentRepairSupport.expectedCommentsAfterInsertColumns(columnIndex, columnCount);
    }
    rejectAffectedColumnStructuresForInsert(sheet, columnIndex);
    Map<Integer, CTCol> shiftedExplicitColumns =
        LayoutAdapter.mutableColumnDefinitionsCopy(
            LayoutAdapter.shiftedForInsert(explicitColumns, columnIndex, columnCount));
    if (columnIndex <= lastColumnIndex) {
      sheet.shiftColumns(columnIndex, lastColumnIndex, columnCount);
      copyAdjacentVisualFormattingIntoInsertedColumns(
          sheet, shiftedExplicitColumns, columnIndex, columnCount, lastColumnIndex);
    }
    LayoutAdapter.rebuildColumnDefinitions(sheet, shiftedExplicitColumns);
    DataValidationAdapter.replaceDataValidations(sheet, expectedValidations);
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  public static void deleteColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      throw new IllegalArgumentException("DELETE_COLUMNS requires at least one existing column");
    }
    if (columns.lastColumnIndex() > lastColumnIndex) {
      throw new IllegalArgumentException(
          "DELETE_COLUMNS columns must stay within existing column bounds: last existing column is "
              + ExcelIndexDisplay.columnValue(lastColumnIndex)
              + "; requested "
              + ExcelIndexDisplay.describe("lastColumnIndex", columns.lastColumnIndex()));
    }
    rejectFormulaBearingWorkbookForColumnEdit(sheet.getWorkbook(), "DELETE_COLUMNS");
    LayoutAdapter.normalizeColumnDefinitionContainer(sheet);
    Map<Integer, CTCol> explicitColumns = LayoutAdapter.snapshotColumnDefinitions(sheet);
    CommentAdapter commentRepairSupport = new CommentAdapter(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<CommentAdapter.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments = commentRepairSupport.expectedCommentsAfterDeleteColumns(columns);
    }
    rejectAffectedColumnStructuresForDelete(sheet, columns);
    if (columns.lastColumnIndex() < lastColumnIndex) {
      sheet.shiftColumns(columns.lastColumnIndex() + 1, lastColumnIndex, -columns.count());
    }
    int clearStart = Math.max(columns.firstColumnIndex(), lastColumnIndex - columns.count() + 1);
    LayoutAdapter.clearTrailingCells(sheet, clearStart, lastColumnIndex);
    LayoutAdapter.rebuildColumnDefinitions(
        sheet, LayoutAdapter.shiftedForDelete(explicitColumns, columns));
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  public static void shiftColumns(XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    requireShiftedColumnBounds(columns, delta);
    rejectFormulaBearingWorkbookForColumnEdit(sheet.getWorkbook(), "SHIFT_COLUMNS");
    LayoutAdapter.normalizeColumnDefinitionContainer(sheet);
    Map<Integer, CTCol> explicitColumns = LayoutAdapter.snapshotColumnDefinitions(sheet);
    CommentAdapter commentRepairSupport = new CommentAdapter(sheet);
    boolean repairComments = commentRepairSupport.hasPersistedComments();
    List<CommentAdapter.CommentRewriteSnapshot> expectedComments = List.of();
    if (repairComments) {
      expectedComments = commentRepairSupport.expectedCommentsAfterShiftColumns(columns, delta);
    }
    rejectAffectedColumnStructuresForShift(sheet, columns, delta);
    sheet.shiftColumns(columns.firstColumnIndex(), columns.lastColumnIndex(), delta);
    LayoutAdapter.rebuildColumnDefinitions(
        sheet, LayoutAdapter.shiftedForShift(explicitColumns, columns, delta));
    if (repairComments) {
      commentRepairSupport.replaceComments(expectedComments);
    }
  }

  public static void setColumnVisibility(XSSFSheet sheet, ExcelColumnSpan columns, boolean hidden) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    LayoutAdapter.normalizeColumnDefinitionContainer(sheet);
    for (int columnIndex = columns.firstColumnIndex();
        columnIndex <= columns.lastColumnIndex();
        columnIndex++) {
      sheet.setColumnHidden(columnIndex, hidden);
    }
    LayoutAdapter.canonicalizeColumnDefinitions(sheet);
  }

  public static void groupColumns(XSSFSheet sheet, ExcelColumnSpan columns, boolean collapsed) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    LayoutAdapter.normalizeColumnDefinitionContainer(sheet);
    sheet.groupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    if (collapsed) {
      LayoutAdapter.collapseColumns(sheet, columns);
      LayoutAdapter.canonicalizeColumnDefinitions(sheet);
      return;
    }
    LayoutAdapter.clearExpandedGroupControlColumn(sheet, columns);
    LayoutAdapter.canonicalizeColumnDefinitions(sheet);
  }

  public static void ungroupColumns(XSSFSheet sheet, ExcelColumnSpan columns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    LayoutAdapter.normalizeColumnDefinitionContainer(sheet);
    LayoutAdapter.prepareColumnsForUngroup(sheet, columns);
    sheet.ungroupColumn(columns.firstColumnIndex(), columns.lastColumnIndex());
    LayoutAdapter.canonicalizeColumnDefinitions(sheet);
  }

  public static int lastColumnIndex(XSSFSheet sheet) {
    return LayoutAdapter.lastColumnIndex(sheet);
  }

  public static List<ColumnLayoutReport> columnLayouts(XSSFSheet sheet) {
    return LayoutAdapter.columnLayouts(sheet);
  }

  private static void requireInsertColumnBounds(XSSFSheet sheet, int columnIndex, int columnCount) {
    int lastColumnIndex = lastColumnIndex(sheet);
    if (columnIndex > lastColumnIndex + 1) {
      throw new IllegalArgumentException(
          "INSERT_COLUMNS "
              + ExcelIndexDisplay.describe("columnIndex", columnIndex)
              + " must be less than or equal to last existing column + 1: "
              + ExcelIndexDisplay.columnValue(lastColumnIndex + 1));
    }
    if (columnIndex + columnCount - 1 > Limit.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_COLUMNS would exceed the maximum column index: destination last column would be "
              + ExcelIndexDisplay.columnValue(columnIndex + columnCount - 1)
              + "; maximum is "
              + ExcelIndexDisplay.columnValue(Limit.MAX_COLUMN_INDEX));
    }
  }

  private static void requireShiftedColumnBounds(ExcelColumnSpan columns, int delta) {
    if (columns.firstColumnIndex() + delta < 0) {
      throw new IllegalArgumentException(
          "SHIFT_COLUMNS would move "
              + ExcelIndexDisplay.describe("firstColumnIndex", columns.firstColumnIndex())
              + " by delta "
              + delta
              + " before the first worksheet column ("
              + ExcelIndexDisplay.excelColumn(0)
              + ")");
    }
    if (columns.lastColumnIndex() + delta > Limit.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "SHIFT_COLUMNS would move "
              + ExcelIndexDisplay.describe("lastColumnIndex", columns.lastColumnIndex())
              + " by delta "
              + delta
              + " beyond the maximum column "
              + ExcelIndexDisplay.columnValue(Limit.MAX_COLUMN_INDEX));
    }
  }

  public static void rejectAffectedRowStructuresForInsert(XSSFSheet sheet, int rowIndex) {
    for (XSSFTable table : sheet.getTables()) {
      CellBounds range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastRow() >= rowIndex) {
        throw unsupportedStructure(
            "INSERT_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move tables are not supported");
      }
    }
    Optional<CellBounds> autofilter = sheetAutofilterRange(sheet);
    if (autofilter.isPresent() && autofilter.orElseThrow().lastRow() >= rowIndex) {
      throw unsupportedStructure(
          "INSERT_ROWS",
          sheet,
          "sheet autofilter " + SheetStructure.formatRange(autofilter.orElseThrow()),
          "row structural edits that would move sheet autofilters are not supported");
    }
  }

  public static void rejectAffectedRowStructuresForDelete(XSSFSheet sheet, ExcelRowSpan rows) {
    for (XSSFTable table : sheet.getTables()) {
      CellBounds range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastRow() >= rows.firstRowIndex()) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move or truncate tables are not supported");
      }
    }
    Optional<CellBounds> autofilter = sheetAutofilterRange(sheet);
    if (autofilter.isPresent() && autofilter.orElseThrow().lastRow() >= rows.firstRowIndex()) {
      throw unsupportedStructure(
          "DELETE_ROWS",
          sheet,
          "sheet autofilter " + SheetStructure.formatRange(autofilter.orElseThrow()),
          "row structural edits that would move or truncate sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      CellBounds range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastRow() >= rows.firstRowIndex()) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            "data validation " + validationRange,
            "row structural edits that would move or truncate data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForRowDelete(sheet.getWorkbook(), sheet, rows);
  }

  public static void rejectAffectedRowStructuresForShift(
      XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    for (XSSFTable table : sheet.getTables()) {
      CellBounds range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (affectsRows(range, rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move tables are not supported");
      }
    }
    Optional<CellBounds> autofilter = sheetAutofilterRange(sheet);
    if (autofilter.isPresent() && affectsRows(autofilter.orElseThrow(), rows, delta)) {
      throw unsupportedStructure(
          "SHIFT_ROWS",
          sheet,
          "sheet autofilter " + SheetStructure.formatRange(autofilter.orElseThrow()),
          "row structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      CellBounds range = parsedRange(validationRange, "data validation", validationRange);
      if (affectsRows(range, rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            "data validation " + validationRange,
            "row structural edits that would move data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForRowShift(sheet.getWorkbook(), sheet, rows, delta);
  }

  public static void rejectAffectedColumnStructuresForInsert(XSSFSheet sheet, int columnIndex) {
    for (XSSFTable table : sheet.getTables()) {
      CellBounds range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastColumn() >= columnIndex) {
        throw unsupportedStructure(
            "INSERT_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move tables are not supported");
      }
    }
    Optional<CellBounds> autofilter = sheetAutofilterRange(sheet);
    if (autofilter.isPresent() && autofilter.orElseThrow().lastColumn() >= columnIndex) {
      throw unsupportedStructure(
          "INSERT_COLUMNS",
          sheet,
          "sheet autofilter " + SheetStructure.formatRange(autofilter.orElseThrow()),
          "column structural edits that would move sheet autofilters are not supported");
    }
  }

  public static void rejectAffectedColumnStructuresForDelete(
      XSSFSheet sheet, ExcelColumnSpan columns) {
    for (XSSFTable table : sheet.getTables()) {
      CellBounds range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastColumn() >= columns.firstColumnIndex()) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move or truncate tables are not supported");
      }
    }
    Optional<CellBounds> autofilter = sheetAutofilterRange(sheet);
    if (autofilter.isPresent()
        && autofilter.orElseThrow().lastColumn() >= columns.firstColumnIndex()) {
      throw unsupportedStructure(
          "DELETE_COLUMNS",
          sheet,
          "sheet autofilter " + SheetStructure.formatRange(autofilter.orElseThrow()),
          "column structural edits that would move or truncate sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      CellBounds range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastColumn() >= columns.firstColumnIndex()) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            "data validation " + validationRange,
            "column structural edits that would move or truncate data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForColumnDelete(sheet.getWorkbook(), sheet, columns);
  }

  public static void rejectAffectedColumnStructuresForShift(
      XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    for (XSSFTable table : sheet.getTables()) {
      CellBounds range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (affectsColumns(range, columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move tables are not supported");
      }
    }
    Optional<CellBounds> autofilter = sheetAutofilterRange(sheet);
    if (autofilter.isPresent() && affectsColumns(autofilter.orElseThrow(), columns, delta)) {
      throw unsupportedStructure(
          "SHIFT_COLUMNS",
          sheet,
          "sheet autofilter " + SheetStructure.formatRange(autofilter.orElseThrow()),
          "column structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      CellBounds range = parsedRange(validationRange, "data validation", validationRange);
      if (affectsColumns(range, columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            "data validation " + validationRange,
            "column structural edits that would move data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForColumnShift(sheet.getWorkbook(), sheet, columns, delta);
  }

  public static void rejectDestructiveNamedRangesForRowDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (overlapsRows(namedRange.range(), rows)) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            namedRange.label(),
            "row structural edits that would truncate range-backed named ranges are not supported");
      }
    }
  }

  public static void rejectDestructiveNamedRangesForRowShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (shiftWouldCorruptRows(namedRange.range(), rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            namedRange.label(),
            "row structural edits that would overwrite or partially move range-backed named ranges are not supported");
      }
    }
  }

  public static void rejectDestructiveNamedRangesForColumnDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (overlapsColumns(namedRange.range(), columns)) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            namedRange.label(),
            "column structural edits that would truncate range-backed named ranges are not supported");
      }
    }
  }

  public static void rejectDestructiveNamedRangesForColumnShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (shiftWouldCorruptColumns(namedRange.range(), columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            namedRange.label(),
            "column structural edits that would overwrite or partially move range-backed named ranges are not supported");
      }
    }
  }

  public static void rejectFormulaBearingWorkbookForColumnEdit(
      XSSFWorkbook workbook, String operationType) {
    if (workbookContainsFormulas(workbook)) {
      throw new IllegalArgumentException(
          operationType
              + " cannot run while workbook formulas are present; formula-bearing column structural edits are unsupported");
    }
    if (workbookContainsFormulaDefinedNames(workbook)) {
      throw new IllegalArgumentException(
          operationType
              + " cannot run while formula-defined names are present; formula-bearing column structural edits are unsupported");
    }
  }

  private static boolean workbookContainsFormulas(XSSFWorkbook workbook) {
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      for (Row row : workbook.getSheetAt(sheetIndex)) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            return true;
          }
        }
      }
    }
    return false;
  }

  public static boolean workbookContainsFormulaDefinedNames(XSSFWorkbook workbook) {
    return workbookContainsFormulaDefinedNames(workbook, workbook.getAllNames());
  }

  public static boolean workbookContainsFormulaDefinedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    for (Name name : names) {
      String refersToFormula = name.getRefersToFormula();
      if (refersToFormula == null || refersToFormula.isBlank()) {
        continue;
      }
      NamedRangeScopeReport scope = scopeForName(workbook, name);
      if (NamedRangeAdapter.resolveTarget(refersToFormula, scope).isEmpty()) {
        return true;
      }
    }
    return false;
  }

  public static Optional<NamedRangeTarget> resolvedRangeBackedTarget(
      XSSFWorkbook workbook, Name name) {
    String refersToFormula = name.getRefersToFormula();
    if (refersToFormula == null || refersToFormula.isBlank()) {
      return Optional.empty();
    }
    return NamedRangeAdapter.resolveTarget(refersToFormula, scopeForName(workbook, name));
  }

  public static List<ResolvedNamedRange> resolvedRangeBackedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    List<ResolvedNamedRange> resolved = new ArrayList<>();
    for (Name name : names) {
      resolvedRangeBackedTarget(workbook, name)
          .ifPresent(
              target ->
                  resolved.add(
                      new ResolvedNamedRange(
                          name.getNameName(),
                          target,
                          CellBounds.parse(((NamedRangeTarget.Range) target).range().value()))));
    }
    return List.copyOf(resolved);
  }

  private static NamedRangeScopeReport scopeForName(XSSFWorkbook workbook, Name name) {
    if (name.getSheetIndex() < 0) {
      return new NamedRangeScopeReport.Workbook();
    }
    return new NamedRangeScopeReport.Sheet(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(
            workbook.getSheetName(name.getSheetIndex())));
  }

  private static IllegalArgumentException unsupportedStructure(
      String operationType, XSSFSheet sheet, String structureLabel, String reason) {
    return new IllegalArgumentException(
        operationType
            + " cannot move "
            + structureLabel
            + " on sheet '"
            + sheet.getSheetName()
            + "'; "
            + reason);
  }

  public static boolean affectsRows(CellBounds range, ExcelRowSpan rows, int delta) {
    int affectedFirstRow = Math.min(rows.firstRowIndex(), rows.firstRowIndex() + delta);
    int affectedLastRow = Math.max(rows.lastRowIndex(), rows.lastRowIndex() + delta);
    return range.firstRow() <= affectedLastRow && range.lastRow() >= affectedFirstRow;
  }

  public static boolean affectsColumns(CellBounds range, ExcelColumnSpan columns, int delta) {
    int affectedFirstColumn =
        Math.min(columns.firstColumnIndex(), columns.firstColumnIndex() + delta);
    int affectedLastColumn = Math.max(columns.lastColumnIndex(), columns.lastColumnIndex() + delta);
    return range.firstColumn() <= affectedLastColumn && range.lastColumn() >= affectedFirstColumn;
  }

  private static boolean overlapsRows(CellBounds range, ExcelRowSpan rows) {
    return range.firstRow() <= rows.lastRowIndex() && range.lastRow() >= rows.firstRowIndex();
  }

  private static boolean overlapsColumns(CellBounds range, ExcelColumnSpan columns) {
    return range.firstColumn() <= columns.lastColumnIndex()
        && range.lastColumn() >= columns.firstColumnIndex();
  }

  private static boolean fullyWithinRows(CellBounds range, ExcelRowSpan rows) {
    return range.firstRow() >= rows.firstRowIndex() && range.lastRow() <= rows.lastRowIndex();
  }

  private static boolean fullyWithinColumns(CellBounds range, ExcelColumnSpan columns) {
    return range.firstColumn() >= columns.firstColumnIndex()
        && range.lastColumn() <= columns.lastColumnIndex();
  }

  public static boolean shiftWouldCorruptRows(CellBounds range, ExcelRowSpan rows, int delta) {
    if (fullyWithinRows(range, rows)) {
      return false;
    }
    ExcelRowSpan destination =
        new ExcelRowSpan(rows.firstRowIndex() + delta, rows.lastRowIndex() + delta);
    return overlapsRows(range, rows) || overlapsRows(range, destination);
  }

  public static boolean shiftWouldCorruptColumns(
      CellBounds range, ExcelColumnSpan columns, int delta) {
    if (fullyWithinColumns(range, columns)) {
      return false;
    }
    ExcelColumnSpan destination =
        new ExcelColumnSpan(columns.firstColumnIndex() + delta, columns.lastColumnIndex() + delta);
    return overlapsColumns(range, columns) || overlapsColumns(range, destination);
  }

  public record ResolvedNamedRange(String name, NamedRangeTarget target, CellBounds range) {
    public ResolvedNamedRange {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(range, "range must not be null");
    }

    private boolean targetsSheet(String sheetName) {
      return ((NamedRangeTarget.Range) target).sheetName().value().equals(sheetName);
    }

    private String label() {
      return "named range '" + name + "'";
    }
  }

  private static Optional<CellBounds> sheetAutofilterRange(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return Optional.empty();
    }
    return Optional.of(
        parsedRange(
            sheet.getCTWorksheet().getAutoFilter().getRef(),
            "sheet autofilter",
            sheet.getCTWorksheet().getAutoFilter().getRef()));
  }

  private static List<String> dataValidationRanges(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    if (dataValidations == null) {
      return List.of();
    }
    List<String> ranges = new ArrayList<>();
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      ranges.addAll(SqrefRanges.normalizedSqref(validation.getSqref()));
    }
    return List.copyOf(ranges);
  }

  private static CellBounds parsedRange(String rawRange, String structureType, String detail) {
    Optional<CellBounds> range = SheetStructure.parseOptionalRange(rawRange);
    if (range.isEmpty()) {
      throw new IllegalArgumentException(
          "Stored " + structureType + " range is invalid and cannot be normalized: " + detail);
    }
    return range.orElseThrow();
  }

  public static void copyAdjacentVisualFormattingIntoInsertedRows(
      XSSFSheet sheet, int rowIndex, int rowCount, int lastRowIndexBeforeShift) {
    Row templateRow = visualTemplateRowForInsert(sheet, rowIndex, lastRowIndexBeforeShift);
    for (int insertedRowIndex = rowIndex;
        insertedRowIndex < rowIndex + rowCount;
        insertedRowIndex++) {
      XSSFRow insertedRow = sheet.createRow(insertedRowIndex);
      copyRowVisualFormatting(templateRow, insertedRow);
    }
  }

  public static void copyAdjacentVisualFormattingIntoInsertedColumns(
      XSSFSheet sheet,
      Map<Integer, CTCol> explicitColumnsAfterShift,
      int columnIndex,
      int columnCount,
      int lastColumnIndexBeforeShift) {
    int templateColumnIndex =
        visualTemplateColumnForInsert(
            sheet, explicitColumnsAfterShift, columnIndex, columnCount, lastColumnIndexBeforeShift);
    CTCol templateColumnDefinition = explicitColumnsAfterShift.get(templateColumnIndex);
    if (templateColumnDefinition != null) {
      for (int insertedColumnIndex = columnIndex;
          insertedColumnIndex < columnIndex + columnCount;
          insertedColumnIndex++) {
        CTCol insertedColumnDefinition = CTCol.Factory.newInstance();
        insertedColumnDefinition.set(templateColumnDefinition);
        explicitColumnsAfterShift.put(insertedColumnIndex, insertedColumnDefinition);
      }
    }
    for (Row row : sheet) {
      copyColumnCellVisualFormatting(row, templateColumnIndex, columnIndex, columnCount);
    }
  }

  private static Row visualTemplateRowForInsert(
      XSSFSheet sheet, int rowIndex, int lastRowIndexBeforeShift) {
    Row nearestAbove = null;
    Row nearestBelow = null;
    for (Row candidateRow : sheet) {
      if (candidateRow.getRowNum() < rowIndex) {
        nearestAbove = candidateRow;
      } else if (nearestBelow == null) {
        nearestBelow = candidateRow;
      }
    }
    return Objects.requireNonNull(
        nearestAbove != null ? nearestAbove : nearestBelow,
        "INSERT_ROWS expected a physical template row after shifting row "
            + ExcelIndexDisplay.rowValue(rowIndex)
            + " inside non-empty sheet bounds ending at "
            + ExcelIndexDisplay.rowValue(lastRowIndexBeforeShift));
  }

  private static void copyRowVisualFormatting(Row templateRow, XSSFRow insertedRow) {
    insertedRow.setHeight(templateRow.getHeight());
    if (templateRow.isFormatted()) {
      insertedRow.setRowStyle(templateRow.getRowStyle());
    }
    for (Cell templateCell : templateRow) {
      Cell insertedCell = insertedRow.createCell(templateCell.getColumnIndex(), CellType.BLANK);
      insertedCell.setCellStyle(templateCell.getCellStyle());
    }
  }

  private static int visualTemplateColumnForInsert(
      XSSFSheet sheet,
      Map<Integer, CTCol> explicitColumnsAfterShift,
      int columnIndex,
      int columnCount,
      int lastColumnIndexBeforeShift) {
    BitSet visualColumns = new BitSet(lastColumnIndexBeforeShift + columnCount + 1);
    explicitColumnsAfterShift.keySet().forEach(visualColumns::set);
    for (Row row : sheet) {
      for (Cell cell : row) {
        visualColumns.set(cell.getColumnIndex());
      }
    }
    int nearestLeft = visualColumns.previousSetBit(columnIndex - 1);
    int templateColumnIndex =
        nearestLeft >= 0 ? nearestLeft : visualColumns.nextSetBit(columnIndex);
    return Objects.checkIndex(templateColumnIndex, lastColumnIndexBeforeShift + columnCount + 1);
  }

  private static void copyColumnCellVisualFormatting(
      Row row, int templateColumnIndex, int columnIndex, int columnCount) {
    Cell templateCell = row.getCell(templateColumnIndex);
    if (templateCell == null) {
      return;
    }
    for (int insertedColumnIndex = columnIndex;
        insertedColumnIndex < columnIndex + columnCount;
        insertedColumnIndex++) {
      Cell insertedCell = row.createCell(insertedColumnIndex, CellType.BLANK);
      insertedCell.setCellStyle(templateCell.getCellStyle());
    }
  }
}
