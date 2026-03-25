package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

/** High-level sheet wrapper for typed reads, writes, and previews. */
public final class ExcelSheet {
  private final Sheet sheet;
  private final WorkbookStyleRegistry styleRegistry;
  private final FormulaEvaluator formulaEvaluator;
  private final DataFormatter dataFormatter;

  ExcelSheet(Sheet sheet, WorkbookStyleRegistry styleRegistry, FormulaEvaluator formulaEvaluator) {
    this.sheet = sheet;
    this.styleRegistry = styleRegistry;
    this.formulaEvaluator = formulaEvaluator;
    this.dataFormatter = new DataFormatter();
  }

  /** Returns the sheet name as defined in the workbook. */
  public String name() {
    return sheet.getSheetName();
  }

  /** Writes a typed value to an A1-style address. */
  public ExcelSheet setCell(String address, ExcelCellValue value) {
    requireNonBlank(address, "address");
    Objects.requireNonNull(value, "value must not be null");

    CellReference cellReference = parseCellReference(address);
    setCell(cellReference.getRow(), cellReference.getCol(), value);
    return this;
  }

  /** Writes a rectangular matrix of values to an A1-style range such as {@code A1:C3}. */
  public ExcelSheet setRange(String range, List<List<ExcelCellValue>> rows) {
    requireNonBlank(range, "range");
    Objects.requireNonNull(rows, "rows must not be null");
    if (rows.isEmpty()) {
      throw new IllegalArgumentException("rows must not be empty");
    }

    List<List<ExcelCellValue>> copiedRows = copyRows(rows);
    ExcelRange excelRange = parseRange(range);
    if (copiedRows.size() != excelRange.rowCount()
        || copiedRows.getFirst().size() != excelRange.columnCount()) {
      throw new IllegalArgumentException(
          "range dimensions do not match provided values: "
              + range
              + " expects "
              + excelRange.rowCount()
              + "x"
              + excelRange.columnCount()
              + " but received "
              + copiedRows.size()
              + "x"
              + copiedRows.getFirst().size());
    }

    for (int rowOffset = 0; rowOffset < copiedRows.size(); rowOffset++) {
      List<ExcelCellValue> rowValues = copiedRows.get(rowOffset);
      for (int columnOffset = 0; columnOffset < rowValues.size(); columnOffset++) {
        setCell(
            excelRange.firstRow() + rowOffset,
            excelRange.firstColumn() + columnOffset,
            rowValues.get(columnOffset));
      }
    }
    return this;
  }

  /** Clears both contents and formatting in a rectangular A1-style range. */
  public ExcelSheet clearRange(String range) {
    requireNonBlank(range, "range");
    ExcelRange excelRange = parseRange(range);
    for (int rowIndex = excelRange.firstRow(); rowIndex <= excelRange.lastRow(); rowIndex++) {
      for (int columnIndex = excelRange.firstColumn();
          columnIndex <= excelRange.lastColumn();
          columnIndex++) {
        Cell cell = getOrCreateCell(rowIndex, columnIndex);
        resetToDefaultStyle(cell);
        cell.setBlank();
      }
    }
    return this;
  }

  /** Applies a style patch to every cell in a rectangular A1-style range. */
  public ExcelSheet applyStyle(String range, ExcelCellStyle style) {
    requireNonBlank(range, "range");
    Objects.requireNonNull(style, "style must not be null");

    ExcelRange excelRange = parseRange(range);
    for (int rowIndex = excelRange.firstRow(); rowIndex <= excelRange.lastRow(); rowIndex++) {
      for (int columnIndex = excelRange.firstColumn();
          columnIndex <= excelRange.lastColumn();
          columnIndex++) {
        Cell cell = getOrCreateCell(rowIndex, columnIndex);
        cell.setCellStyle(styleRegistry.mergedStyle(cell, style));
      }
    }
    return this;
  }

  /** Appends a new row using the next available row index. */
  public ExcelSheet appendRow(ExcelCellValue... values) {
    Objects.requireNonNull(values, "values must not be null");
    for (ExcelCellValue value : values) {
      Objects.requireNonNull(value, "values must not contain nulls");
    }

    int rowIndex = sheet.getPhysicalNumberOfRows() == 0 ? 0 : sheet.getLastRowNum() + 1;
    for (int columnIndex = 0; columnIndex < values.length; columnIndex++) {
      setCell(rowIndex, columnIndex, values[columnIndex]);
    }

    return this;
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  public ExcelSheet mergeCells(String range) {
    requireNonBlank(range, "range");

    ExcelRange excelRange = parseRange(range);
    requireMergeableRange(range, excelRange);
    if (findMergedRegionIndex(sheet, excelRange) >= 0) {
      return this;
    }
    requireNoMergedRegionOverlap(sheet, excelRange);
    sheet.addMergedRegion(toCellRangeAddress(excelRange));
    return this;
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  public ExcelSheet unmergeCells(String range) {
    requireNonBlank(range, "range");

    ExcelRange excelRange = parseRange(range);
    int mergedRegionIndex = findMergedRegionIndex(sheet, excelRange);
    if (mergedRegionIndex < 0) {
      throw new IllegalArgumentException("No merged region matches range: " + range);
    }
    sheet.removeMergedRegion(mergedRegionIndex);
    return this;
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  public ExcelSheet setColumnWidth(
      int firstColumnIndex, int lastColumnIndex, double widthCharacters) {
    requireNonNegative(firstColumnIndex, "firstColumnIndex");
    requireNonNegative(lastColumnIndex, "lastColumnIndex");
    requireOrderedSpan(firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");

    int widthUnits = toColumnWidthUnits(widthCharacters);
    for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
      sheet.setColumnWidth(columnIndex, widthUnits);
    }
    return this;
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  public ExcelSheet setRowHeight(int firstRowIndex, int lastRowIndex, double heightPoints) {
    requireNonNegative(firstRowIndex, "firstRowIndex");
    requireNonNegative(lastRowIndex, "lastRowIndex");
    requireOrderedSpan(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");

    float heightPointsValue = toRowHeightPoints(heightPoints);
    for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
      getOrCreateRow(rowIndex).setHeightInPoints(heightPointsValue);
    }
    return this;
  }

  /** Freezes panes using explicit split and visible-origin coordinates. */
  public ExcelSheet freezePanes(int splitColumn, int splitRow, int leftmostColumn, int topRow) {
    requireNonNegative(splitColumn, "splitColumn");
    requireNonNegative(splitRow, "splitRow");
    requireNonNegative(leftmostColumn, "leftmostColumn");
    requireNonNegative(topRow, "topRow");
    requireFreezePaneCoordinates(splitColumn, splitRow, leftmostColumn, topRow);

    sheet.createFreezePane(splitColumn, splitRow, leftmostColumn, topRow);
    return this;
  }

  /** Returns the number of physically stored rows in the sheet. */
  public int physicalRowCount() {
    return sheet.getPhysicalNumberOfRows();
  }

  /** Returns the 0-based index of the last row, or -1 if no rows exist. */
  public int lastRowIndex() {
    return sheet.getLastRowNum();
  }

  /**
   * Returns the 0-based index of the widest column across all rows, or -1 if the sheet is empty.
   */
  public int lastColumnIndex() {
    int lastColumnIndex = -1;
    for (Row row : sheet) {
      lastColumnIndex = Math.max(lastColumnIndex, row.getLastCellNum() - 1);
    }
    return lastColumnIndex;
  }

  /** Auto-sizes all populated columns on this sheet to fit their content. */
  public ExcelSheet autoSizeColumns() {
    int lastColumn = lastColumnIndex();
    for (int col = 0; col <= lastColumn; col++) {
      sheet.autoSizeColumn(col);
    }
    return this;
  }

  /** Reads a string cell by A1-style address. */
  public String text(String address) {
    requireNonBlank(address, "address");
    return requiredCell(address).getStringCellValue();
  }

  /** Reads a numeric cell by A1-style address, evaluating formulas when needed. */
  public double number(String address) {
    requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      org.apache.poi.ss.usermodel.CellValue evaluatedCell;
      try {
        evaluatedCell = formulaEvaluator.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(name(), address, cell.getCellFormula(), exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.NUMERIC) {
        throw new IllegalStateException("Cell does not evaluate to a numeric value: " + address);
      }
      return evaluatedCell.getNumberValue();
    }
    return cell.getNumericCellValue();
  }

  /** Reads a boolean cell by A1-style address, evaluating formulas when needed. */
  public boolean bool(String address) {
    requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      org.apache.poi.ss.usermodel.CellValue evaluatedCell;
      try {
        evaluatedCell = formulaEvaluator.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(name(), address, cell.getCellFormula(), exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.BOOLEAN) {
        throw new IllegalStateException("Cell does not evaluate to a boolean value: " + address);
      }
      return evaluatedCell.getBooleanValue();
    }
    return cell.getBooleanCellValue();
  }

  /** Reads the raw formula expression stored in a formula cell. */
  public String formula(String address) {
    requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() != CellType.FORMULA) {
      throw new IllegalStateException("Cell does not contain a formula: " + address);
    }
    return cell.getCellFormula();
  }

  /** Captures a formatted, typed snapshot of a single cell. */
  public ExcelCellSnapshot snapshotCell(String address) {
    requireNonBlank(address, "address");
    return snapshot(address, requiredCell(address));
  }

  /** Returns a compact preview of the top-left portion of the sheet. */
  public List<ExcelPreviewRow> preview(int maxRows, int maxColumns) {
    if (maxRows <= 0) {
      throw new IllegalArgumentException("maxRows must be greater than 0");
    }
    if (maxColumns <= 0) {
      throw new IllegalArgumentException("maxColumns must be greater than 0");
    }

    List<ExcelPreviewRow> previewRows = new ArrayList<>();
    int endingRowIndex = Math.min(lastRowIndex(), maxRows - 1);
    for (int rowIndex = 0; rowIndex <= endingRowIndex; rowIndex++) {
      previewRows.add(new ExcelPreviewRow(rowIndex, previewRow(rowIndex, maxColumns)));
    }
    return List.copyOf(previewRows);
  }

  private List<ExcelCellSnapshot> previewRow(int rowIndex, int maxColumns) {
    Row row = sheet.getRow(rowIndex);
    List<ExcelCellSnapshot> cells = new ArrayList<>();
    if (row != null) {
      for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (shouldPreview(cell)) {
          cells.add(snapshot(new CellReference(rowIndex, columnIndex).formatAsString(), cell));
        }
      }
    }
    return List.copyOf(cells);
  }

  private void setCell(int rowIndex, int columnIndex, ExcelCellValue value) {
    Cell cell = getOrCreateCell(rowIndex, columnIndex);

    switch (value) {
      case ExcelCellValue.BlankValue _ -> {
        resetToDefaultStyle(cell);
        cell.setBlank();
      }
      case ExcelCellValue.TextValue textValue -> {
        resetToDefaultStyle(cell);
        cell.setCellValue(textValue.value());
      }
      case ExcelCellValue.NumberValue numberValue -> {
        resetToDefaultStyle(cell);
        cell.setCellValue(numberValue.value());
      }
      case ExcelCellValue.BooleanValue booleanValue -> {
        resetToDefaultStyle(cell);
        cell.setCellValue(booleanValue.value());
      }
      case ExcelCellValue.DateValue dateValue -> {
        cell.setCellValue(dateValue.value());
        cell.setCellStyle(styleRegistry.localDateStyle());
      }
      case ExcelCellValue.DateTimeValue dateTimeValue -> {
        cell.setCellValue(dateTimeValue.value());
        cell.setCellStyle(styleRegistry.localDateTimeStyle());
      }
      case ExcelCellValue.FormulaValue formulaValue -> {
        resetToDefaultStyle(cell);
        try {
          cell.setCellFormula(formulaValue.expression());
        } catch (RuntimeException exception) {
          throw FormulaExceptions.wrap(
              name(),
              new CellReference(rowIndex, columnIndex).formatAsString(),
              formulaValue.expression(),
              exception);
        }
      }
    }
  }

  private Cell requiredCell(String address) {
    CellReference cellReference = parseCellReference(address);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }

    Cell cell = row.getCell(cellReference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }

    return cell;
  }

  private Cell getOrCreateCell(int rowIndex, int columnIndex) {
    return getOrCreateRow(rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
  }

  private Row getOrCreateRow(int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  private ExcelCellSnapshot snapshot(String address, Cell cell) {
    CellType declaredType = cell.getCellType();
    String formulaExpression = declaredType == CellType.FORMULA ? cell.getCellFormula() : null;
    String displayValue;
    try {
      displayValue =
          declaredType == CellType.FORMULA
              ? dataFormatter.formatCellValue(cell, formulaEvaluator)
              : dataFormatter.formatCellValue(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(name(), address, formulaExpression, exception);
    }
    ExcelCellStyleSnapshot style = styleRegistry.snapshot(cell);

    if (declaredType == CellType.FORMULA) {
      String formula = cell.getCellFormula();
      CellValue evaluatedCell;
      try {
        evaluatedCell = formulaEvaluator.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(name(), address, formula, exception);
      }

      CellType evalType = evaluatedCell != null ? evaluatedCell.getCellType() : CellType.BLANK;
      ExcelCellSnapshot evaluation =
          switch (evalType) {
            case STRING ->
                new ExcelCellSnapshot.TextSnapshot(
                    address, "STRING", displayValue, style, evaluatedCell.getStringValue());
            case NUMERIC ->
                new ExcelCellSnapshot.NumberSnapshot(
                    address, "NUMERIC", displayValue, style, evaluatedCell.getNumberValue());
            case BOOLEAN ->
                new ExcelCellSnapshot.BooleanSnapshot(
                    address, "BOOLEAN", displayValue, style, evaluatedCell.getBooleanValue());
            case ERROR ->
                new ExcelCellSnapshot.ErrorSnapshot(
                    address,
                    "ERROR",
                    displayValue,
                    style,
                    FormulaError.forInt(evaluatedCell.getErrorValue()).getString());
            default -> new ExcelCellSnapshot.BlankSnapshot(address, "BLANK", displayValue, style);
          };
      return new ExcelCellSnapshot.FormulaSnapshot(
          address, "FORMULA", displayValue, style, formula, evaluation);
    }

    return switch (declaredType) {
      case STRING ->
          new ExcelCellSnapshot.TextSnapshot(
              address, "STRING", displayValue, style, cell.getStringCellValue());
      case NUMERIC ->
          new ExcelCellSnapshot.NumberSnapshot(
              address, "NUMERIC", displayValue, style, cell.getNumericCellValue());
      case BOOLEAN ->
          new ExcelCellSnapshot.BooleanSnapshot(
              address, "BOOLEAN", displayValue, style, cell.getBooleanCellValue());
      case ERROR ->
          new ExcelCellSnapshot.ErrorSnapshot(
              address,
              "ERROR",
              displayValue,
              style,
              FormulaError.forInt(cell.getErrorCellValue()).getString());
      case BLANK, _NONE, FORMULA ->
          new ExcelCellSnapshot.BlankSnapshot(address, "BLANK", displayValue, style);
    };
  }

  private void resetToDefaultStyle(Cell cell) {
    cell.setCellStyle(styleRegistry.defaultStyle());
  }

  private CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  private ExcelRange parseRange(String range) {
    return ExcelRange.parse(range);
  }

  private void requireMergeableRange(String range, ExcelRange excelRange) {
    if (excelRange.rowCount() == 1 && excelRange.columnCount() == 1) {
      throw new IllegalArgumentException("range must span at least two cells: " + range);
    }
  }

  /** Rejects merges that would overlap any existing merged region on the sheet. */
  static void requireNoMergedRegionOverlap(Sheet sheet, ExcelRange excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      CellRangeAddress existing = sheet.getMergedRegion(regionIndex);
      if (intersects(existing, excelRange)) {
        throw new IllegalArgumentException(
            "Merged range overlaps existing merged region: " + existing.formatAsString());
      }
    }
  }

  /** Returns the exact merged-region index for the range, or {@code -1} when none matches. */
  static int findMergedRegionIndex(Sheet sheet, ExcelRange excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      if (matches(sheet.getMergedRegion(regionIndex), excelRange)) {
        return regionIndex;
      }
    }
    return -1;
  }

  private CellRangeAddress toCellRangeAddress(ExcelRange excelRange) {
    return new CellRangeAddress(
        excelRange.firstRow(),
        excelRange.lastRow(),
        excelRange.firstColumn(),
        excelRange.lastColumn());
  }

  /** Returns whether the POI range exactly matches the GridGrind range coordinates. */
  static boolean matches(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return rangeAddress.getFirstRow() == excelRange.firstRow()
        && rangeAddress.getLastRow() == excelRange.lastRow()
        && rangeAddress.getFirstColumn() == excelRange.firstColumn()
        && rangeAddress.getLastColumn() == excelRange.lastColumn();
  }

  /** Returns whether the POI range intersects the GridGrind range at any cell. */
  static boolean intersects(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return rangeAddress.getFirstRow() <= excelRange.lastRow()
        && rangeAddress.getLastRow() >= excelRange.firstRow()
        && rangeAddress.getFirstColumn() <= excelRange.lastColumn()
        && rangeAddress.getLastColumn() >= excelRange.firstColumn();
  }

  /** Converts Excel character-width input into POI column-width units. */
  static int toColumnWidthUnits(double widthCharacters) {
    requireFinitePositive(widthCharacters, "widthCharacters");
    if (widthCharacters > 255.0d) {
      throw new IllegalArgumentException(
          "widthCharacters must be less than or equal to 255.0: " + widthCharacters);
    }
    int widthUnits = (int) Math.round(widthCharacters * 256.0d);
    if (widthUnits <= 0) {
      throw new IllegalArgumentException(
          "widthCharacters is too small to produce a visible Excel column width: "
              + widthCharacters);
    }
    return widthUnits;
  }

  /** Converts row-height input to POI points after validating Excel storage constraints. */
  static float toRowHeightPoints(double heightPoints) {
    requireFinitePositive(heightPoints, "heightPoints");
    long heightTwips = Math.round(heightPoints * 20.0d);
    if (heightTwips > Short.MAX_VALUE) {
      throw new IllegalArgumentException(
          "heightPoints is too large for Excel row height storage: " + heightPoints);
    }
    if (heightTwips <= 0L) {
      throw new IllegalArgumentException(
          "heightPoints is too small to produce a visible Excel row height: " + heightPoints);
    }
    return (float) heightPoints;
  }

  /** Validates freeze-pane split and visible-origin coordinates before POI mutation. */
  static void requireFreezePaneCoordinates(
      int splitColumn, int splitRow, int leftmostColumn, int topRow) {
    if (splitColumn == 0 && splitRow == 0) {
      throw new IllegalArgumentException("splitColumn and splitRow must not both be 0");
    }
    if (splitColumn == 0 && leftmostColumn != 0) {
      throw new IllegalArgumentException(
          "leftmostColumn must be 0 when splitColumn is 0: " + leftmostColumn);
    }
    if (splitRow == 0 && topRow != 0) {
      throw new IllegalArgumentException("topRow must be 0 when splitRow is 0: " + topRow);
    }
    if (splitColumn > 0 && leftmostColumn < splitColumn) {
      throw new IllegalArgumentException(
          "leftmostColumn must be greater than or equal to splitColumn");
    }
    if (splitRow > 0 && topRow < splitRow) {
      throw new IllegalArgumentException("topRow must be greater than or equal to splitRow");
    }
  }

  private boolean shouldPreview(Cell cell) {
    return cell != null
        && (cell.getCellType() != CellType.BLANK || cell.getCellStyle().getIndex() != 0);
  }

  private List<List<ExcelCellValue>> copyRows(List<List<ExcelCellValue>> rows) {
    List<List<ExcelCellValue>> copiedRows = new ArrayList<>(rows.size());
    int expectedWidth = -1;
    for (List<ExcelCellValue> row : rows) {
      Objects.requireNonNull(row, "rows must not contain nulls");
      List<ExcelCellValue> copiedRow = List.copyOf(row);
      if (copiedRow.isEmpty()) {
        throw new IllegalArgumentException("rows must not contain empty rows");
      }
      if (expectedWidth < 0) {
        expectedWidth = copiedRow.size();
      } else if (copiedRow.size() != expectedWidth) {
        throw new IllegalArgumentException("rows must describe a rectangular matrix");
      }
      for (ExcelCellValue value : copiedRow) {
        Objects.requireNonNull(value, "rows must not contain null cell values");
      }
      copiedRows.add(copiedRow);
    }
    return List.copyOf(copiedRows);
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static void requireOrderedSpan(
      int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
    if (lastValue < firstValue) {
      throw new IllegalArgumentException(
          lastFieldName + " must not be less than " + firstFieldName);
    }
  }

  private static void requireFinitePositive(double value, String fieldName) {
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    if (value <= 0.0d) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
  }
}
