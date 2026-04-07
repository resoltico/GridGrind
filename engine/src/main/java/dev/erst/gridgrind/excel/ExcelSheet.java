package dev.erst.gridgrind.excel;

import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** High-level sheet wrapper for typed reads, writes, and previews. */
@SuppressWarnings("PMD.ExcessivePublicCount")
public final class ExcelSheet {
  private final Sheet sheet;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final DataFormatter dataFormatter;
  private final ExcelDataValidationController dataValidationController;
  private final ExcelConditionalFormattingController conditionalFormattingController;
  private final ExcelAutofilterController autofilterController;
  private final ExcelPrintLayoutController printLayoutController;
  private final ExcelRowColumnStructureController rowColumnStructureController;

  ExcelSheet(Sheet sheet, WorkbookStyleRegistry styleRegistry, ExcelFormulaRuntime formulaRuntime) {
    this.sheet = sheet;
    this.styleRegistry = styleRegistry;
    this.formulaRuntime = formulaRuntime;
    this.dataFormatter = new DataFormatter();
    this.dataValidationController = new ExcelDataValidationController();
    this.conditionalFormattingController = new ExcelConditionalFormattingController();
    this.autofilterController = new ExcelAutofilterController();
    this.printLayoutController = new ExcelPrintLayoutController();
    this.rowColumnStructureController = new ExcelRowColumnStructureController();
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  ExcelSheet(Sheet sheet, WorkbookStyleRegistry styleRegistry, FormulaEvaluator formulaEvaluator) {
    this(sheet, styleRegistry, ExcelFormulaRuntime.poi(formulaEvaluator));
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
    writeCellValue(cellReference.getRow(), cellReference.getCol(), value);
    syncTableHeaders(
        new ExcelRange(
            cellReference.getRow(),
            cellReference.getRow(),
            cellReference.getCol(),
            cellReference.getCol()));
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
        writeCellValue(
            excelRange.firstRow() + rowOffset,
            excelRange.firstColumn() + columnOffset,
            rowValues.get(columnOffset));
      }
    }
    syncTableHeaders(excelRange);
    return this;
  }

  /** Clears both contents and formatting in a rectangular A1-style range. */
  public ExcelSheet clearRange(String range) {
    requireNonBlank(range, "range");
    ExcelRange excelRange = parseRange(range);
    for (int rowIndex = excelRange.firstRow(); rowIndex <= excelRange.lastRow(); rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      if (row == null) {
        continue;
      }
      for (int columnIndex = excelRange.firstColumn();
          columnIndex <= excelRange.lastColumn();
          columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
          continue;
        }
        resetToDefaultStyle(cell);
        cell.removeHyperlink();
        cell.removeCellComment();
        cell.setBlank();
      }
    }
    syncTableHeaders(excelRange);
    return this;
  }

  /** Replaces the hyperlink attached to one cell, creating the cell if necessary. */
  public ExcelSheet setHyperlink(String address, ExcelHyperlink hyperlink) {
    requireNonBlank(address, "address");
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");

    CellReference cellReference = parseCellReference(address);
    Cell cell = getOrCreateCell(cellReference.getRow(), cellReference.getCol());
    cell.removeHyperlink();
    org.apache.poi.ss.usermodel.Hyperlink poiHyperlink =
        sheet.getWorkbook().getCreationHelper().createHyperlink(toPoi(hyperlink.type()));
    poiHyperlink.setAddress(toPoiTarget(hyperlink));
    cell.setHyperlink(poiHyperlink);
    return this;
  }

  /** Removes any hyperlink attached to one cell; no-op when the cell does not physically exist. */
  public ExcelSheet clearHyperlink(String address) {
    requireNonBlank(address, "address");
    Cell cell = optionalCell(address);
    if (cell != null) {
      cell.removeHyperlink();
    }
    return this;
  }

  /** Replaces the plain-text comment attached to one cell, creating the cell if necessary. */
  public ExcelSheet setComment(String address, ExcelComment comment) {
    requireNonBlank(address, "address");
    Objects.requireNonNull(comment, "comment must not be null");

    CellReference cellReference = parseCellReference(address);
    Cell cell = getOrCreateCell(cellReference.getRow(), cellReference.getCol());
    cell.setCellComment(newComment(cellReference.getRow(), cellReference.getCol(), comment));
    return this;
  }

  /** Removes any comment attached to one cell; no-op when the cell does not physically exist. */
  public ExcelSheet clearComment(String address) {
    requireNonBlank(address, "address");
    Cell cell = optionalCell(address);
    if (cell != null) {
      cell.removeCellComment();
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
    syncTableHeaders(excelRange);
    return this;
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  public ExcelSheet setDataValidation(String range, ExcelDataValidationDefinition validation) {
    requireNonBlank(range, "range");
    Objects.requireNonNull(validation, "validation must not be null");
    dataValidationController.setDataValidation(xssfSheet(), range, validation);
    return this;
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  public ExcelSheet clearDataValidations(ExcelRangeSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    dataValidationController.clearDataValidations(xssfSheet(), selection);
    return this;
  }

  /** Creates or replaces one logical conditional-formatting block on this sheet. */
  public ExcelSheet setConditionalFormatting(ExcelConditionalFormattingBlockDefinition block) {
    Objects.requireNonNull(block, "block must not be null");
    conditionalFormattingController.setConditionalFormatting(xssfSheet(), block);
    return this;
  }

  /** Removes conditional-formatting blocks on this sheet matching the provided selection. */
  public ExcelSheet clearConditionalFormatting(ExcelRangeSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    conditionalFormattingController.clearConditionalFormatting(xssfSheet(), selection);
    return this;
  }

  /** Creates or replaces one sheet-level autofilter range. */
  public ExcelSheet setAutofilter(String range) {
    requireNonBlank(range, "range");
    autofilterController.setSheetAutofilter(xssfSheet(), range);
    return this;
  }

  /** Clears the sheet-level autofilter range on this sheet. */
  public ExcelSheet clearAutofilter() {
    autofilterController.clearSheetAutofilter(xssfSheet());
    return this;
  }

  /** Appends a new row using the next available row index. */
  public ExcelSheet appendRow(ExcelCellValue... values) {
    Objects.requireNonNull(values, "values must not be null");
    for (ExcelCellValue value : values) {
      Objects.requireNonNull(value, "values must not contain nulls");
    }

    int rowIndex = nextAppendRowIndex();
    for (int columnIndex = 0; columnIndex < values.length; columnIndex++) {
      writeCellValue(rowIndex, columnIndex, values[columnIndex]);
    }
    if (values.length > 0) {
      syncTableHeaders(new ExcelRange(rowIndex, rowIndex, 0, values.length - 1));
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

  /** Inserts one or more blank rows before the provided zero-based row index. */
  public ExcelSheet insertRows(int rowIndex, int rowCount) {
    rowColumnStructureController.insertRows(xssfSheet(), rowIndex, rowCount);
    return this;
  }

  /** Deletes the requested inclusive zero-based row band. */
  public ExcelSheet deleteRows(ExcelRowSpan rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.deleteRows(xssfSheet(), rows);
    return this;
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  public ExcelSheet shiftRows(ExcelRowSpan rows, int delta) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.shiftRows(xssfSheet(), rows, delta);
    return this;
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  public ExcelSheet insertColumns(int columnIndex, int columnCount) {
    rowColumnStructureController.insertColumns(xssfSheet(), columnIndex, columnCount);
    return this;
  }

  /** Deletes the requested inclusive zero-based column band. */
  public ExcelSheet deleteColumns(ExcelColumnSpan columns) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.deleteColumns(xssfSheet(), columns);
    return this;
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  public ExcelSheet shiftColumns(ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.shiftColumns(xssfSheet(), columns, delta);
    return this;
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  public ExcelSheet setRowVisibility(ExcelRowSpan rows, boolean hidden) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.setRowVisibility(xssfSheet(), rows, hidden);
    return this;
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  public ExcelSheet setColumnVisibility(ExcelColumnSpan columns, boolean hidden) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.setColumnVisibility(xssfSheet(), columns, hidden);
    return this;
  }

  /** Applies one outline group to the requested inclusive zero-based row band. */
  public ExcelSheet groupRows(ExcelRowSpan rows, boolean collapsed) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.groupRows(xssfSheet(), rows, collapsed);
    return this;
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  public ExcelSheet ungroupRows(ExcelRowSpan rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.ungroupRows(xssfSheet(), rows);
    return this;
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  public ExcelSheet groupColumns(ExcelColumnSpan columns, boolean collapsed) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.groupColumns(xssfSheet(), columns, collapsed);
    return this;
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  public ExcelSheet ungroupColumns(ExcelColumnSpan columns) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.ungroupColumns(xssfSheet(), columns);
    return this;
  }

  /** Applies one explicit pane state to this sheet. */
  public ExcelSheet setPane(ExcelSheetPane pane) {
    Objects.requireNonNull(pane, "pane must not be null");
    ExcelSheetViewSupport.setPane(xssfSheet(), pane);
    return this;
  }

  /** Applies one explicit zoom percentage to this sheet. */
  public ExcelSheet setZoom(int zoomPercent) {
    ExcelSheetViewSupport.requireZoomPercent(zoomPercent);
    ExcelSheetViewSupport.setZoomPercent(xssfSheet(), zoomPercent);
    return this;
  }

  /** Applies the provided print layout as the authoritative supported print state. */
  public ExcelSheet setPrintLayout(ExcelPrintLayout printLayout) {
    Objects.requireNonNull(printLayout, "printLayout must not be null");
    printLayoutController.setPrintLayout(xssfSheet(), printLayout);
    return this;
  }

  /** Clears the supported print layout state from this sheet. */
  public ExcelSheet clearPrintLayout() {
    printLayoutController.clearPrintLayout(xssfSheet());
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
    return rowColumnStructureController.lastColumnIndex(xssfSheet());
  }

  /** Auto-sizes all populated columns on this sheet to fit their content. */
  public ExcelSheet autoSizeColumns() {
    DeterministicColumnSizer.autoSize(sheet, name(), dataFormatter, formulaRuntime);
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
        evaluatedCell = formulaRuntime.evaluate(cell);
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
        evaluatedCell = formulaRuntime.evaluate(cell);
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

  /** Captures a formatted, typed snapshot of a single cell, returning blank for unwritten cells. */
  public ExcelCellSnapshot snapshotCell(String address) {
    requireNonBlank(address, "address");
    CellReference cellReference = parseCellReference(address);
    requireValidCellReference(address, cellReference);
    return snapshotCellOrBlank(address, cellReference.getRow(), cellReference.getCol());
  }

  /** Captures exact snapshots for the provided ordered A1 addresses. */
  public List<ExcelCellSnapshot> snapshotCells(List<String> addresses) {
    Objects.requireNonNull(addresses, "addresses must not be null");
    List<ExcelCellSnapshot> cells = new ArrayList<>(addresses.size());
    for (String address : List.copyOf(addresses)) {
      requireNonBlank(address, "address");
      cells.add(snapshotCell(address));
    }
    return List.copyOf(cells);
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

  /** Returns an exact rectangular window of cell snapshots anchored at one top-left address. */
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  public WorkbookReadResult.Window window(String topLeftAddress, int rowCount, int columnCount) {
    requireNonBlank(topLeftAddress, "topLeftAddress");
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be greater than 0");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be greater than 0");
    }

    CellReference topLeft = parseCellReference(topLeftAddress);
    requireValidCellReference(topLeftAddress, topLeft);
    int lastRow = topLeft.getRow() + rowCount - 1;
    int lastCol = topLeft.getCol() + columnCount - 1;
    if (lastRow > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || lastCol > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new IllegalArgumentException(
          "window extends beyond the Excel sheet boundary: topLeft="
              + topLeftAddress
              + " rowCount="
              + rowCount
              + " columnCount="
              + columnCount);
    }
    List<WorkbookReadResult.WindowRow> rows = new ArrayList<>(rowCount);
    for (int rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      int rowIndex = topLeft.getRow() + rowOffset;
      List<ExcelCellSnapshot> cells = new ArrayList<>(columnCount);
      for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
        int columnIndex = topLeft.getCol() + columnOffset;
        String address = new CellReference(rowIndex, columnIndex).formatAsString();
        cells.add(snapshotCellOrBlank(address, rowIndex, columnIndex));
      }
      rows.add(new WorkbookReadResult.WindowRow(rowIndex, List.copyOf(cells)));
    }
    return new WorkbookReadResult.Window(
        name(), topLeftAddress, rowCount, columnCount, List.copyOf(rows));
  }

  /** Returns every merged region currently defined on the sheet. */
  public List<WorkbookReadResult.MergedRegion> mergedRegions() {
    List<WorkbookReadResult.MergedRegion> mergedRegions =
        new ArrayList<>(sheet.getNumMergedRegions());
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      mergedRegions.add(
          new WorkbookReadResult.MergedRegion(sheet.getMergedRegion(regionIndex).formatAsString()));
    }
    return List.copyOf(mergedRegions);
  }

  /** Returns hyperlink metadata for the selected cells on this sheet. */
  public List<WorkbookReadResult.CellHyperlink> hyperlinks(ExcelCellSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return switch (selection) {
      case ExcelCellSelection.AllUsedCells _ -> allUsedHyperlinks();
      case ExcelCellSelection.Selected selected -> selectedHyperlinks(selected.addresses());
    };
  }

  /** Returns comment metadata for the selected cells on this sheet. */
  public List<WorkbookReadResult.CellComment> comments(ExcelCellSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return switch (selection) {
      case ExcelCellSelection.AllUsedCells _ -> allUsedComments();
      case ExcelCellSelection.Selected selected -> selectedComments(selected.addresses());
    };
  }

  /** Returns layout metadata such as panes, zoom, and visible sizing. */
  public WorkbookReadResult.SheetLayout layout() {
    return new WorkbookReadResult.SheetLayout(
        name(),
        ExcelSheetViewSupport.pane(xssfSheet()),
        ExcelSheetViewSupport.zoomPercent(xssfSheet()),
        rowColumnStructureController.columnLayouts(xssfSheet()),
        rowColumnStructureController.rowLayouts(xssfSheet()));
  }

  /** Returns supported print-layout metadata for this sheet. */
  public ExcelPrintLayout printLayout() {
    return printLayoutController.printLayout(xssfSheet());
  }

  /** Returns data-validation metadata for the selected ranges on this sheet. */
  public List<ExcelDataValidationSnapshot> dataValidations(ExcelRangeSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return dataValidationController.dataValidations(xssfSheet(), selection);
  }

  /** Returns factual conditional-formatting blocks for the selected ranges on this sheet. */
  public List<ExcelConditionalFormattingBlockSnapshot> conditionalFormatting(
      ExcelRangeSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return conditionalFormattingController.conditionalFormatting(xssfSheet(), selection);
  }

  /** Returns every formula cell currently present on the sheet. */
  public List<ExcelCellSnapshot.FormulaSnapshot> formulaCells() {
    List<ExcelCellSnapshot.FormulaSnapshot> formulas = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          formulas.add(
              (ExcelCellSnapshot.FormulaSnapshot)
                  snapshot(
                      new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                      cell));
        }
      }
    }
    return List.copyOf(formulas);
  }

  /** Returns the number of formula cells currently present on the sheet. */
  int formulaCellCount() {
    int count = 0;
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          count++;
        }
      }
    }
    return count;
  }

  /** Returns the number of conditional-formatting blocks currently present on the sheet. */
  int conditionalFormattingBlockCount() {
    return conditionalFormattingController.conditionalFormattingBlockCount(xssfSheet());
  }

  /** Returns derived findings for formula health on this sheet. */
  List<WorkbookAnalysis.AnalysisFinding> formulaHealthFindings() {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() != CellType.FORMULA) {
          continue;
        }
        String address = cellAddress(cell);
        String formula = cell.getCellFormula();
        WorkbookAnalysis.AnalysisLocation.Cell location = cellLocation(address);
        if (containsExternalWorkbookReference(formula)) {
          findings.add(
              new WorkbookAnalysis.AnalysisFinding(
                  WorkbookAnalysis.AnalysisFindingCode.FORMULA_EXTERNAL_REFERENCE,
                  WorkbookAnalysis.AnalysisSeverity.WARNING,
                  "External workbook reference",
                  "Formula references an external workbook and may not evaluate deterministically.",
                  location,
                  List.of(formula)));
        }
        volatileFunctions(formula)
            .forEach(
                functionName ->
                    findings.add(
                        new WorkbookAnalysis.AnalysisFinding(
                            WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                            WorkbookAnalysis.AnalysisSeverity.INFO,
                            "Volatile formula function",
                            "Formula uses volatile function " + functionName + ".",
                            location,
                            List.of(formula, functionName))));
        try {
          CellValue evaluated = formulaRuntime.evaluate(cell);
          if (evaluated != null && evaluated.getCellType() == CellType.ERROR) {
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    WorkbookAnalysis.AnalysisSeverity.ERROR,
                    "Formula evaluates to an error",
                    "Formula currently evaluates to "
                        + FormulaError.forInt(evaluated.getErrorValue()).getString()
                        + ".",
                    location,
                    List.of(formula)));
          }
        } catch (RuntimeException exception) {
          findings.add(
              new WorkbookAnalysis.AnalysisFinding(
                  WorkbookAnalysis.AnalysisFindingCode.FORMULA_EVALUATION_FAILURE,
                  WorkbookAnalysis.AnalysisSeverity.ERROR,
                  "Formula evaluation failed",
                  "Formula evaluation failed: " + exceptionMessage(exception),
                  location,
                  List.of(formula, exception.getClass().getSimpleName())));
        }
      }
    }
    return List.copyOf(findings);
  }

  /** Returns derived conditional-formatting health findings on this sheet. */
  List<WorkbookAnalysis.AnalysisFinding> conditionalFormattingHealthFindings() {
    return conditionalFormattingController.conditionalFormattingHealthFindings(name(), xssfSheet());
  }

  /** Returns the number of raw hyperlinks currently present on the sheet. */
  int hyperlinkCount() {
    int count = 0;
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (hasUsableHyperlink(cell.getHyperlink())) {
          count++;
        }
      }
    }
    return count;
  }

  /** Returns derived findings for hyperlink health on this sheet. */
  /**
   * Returns derived findings for hyperlink health on this sheet using the workbook path context.
   */
  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings(
      WorkbookLocation workbookLocation) {
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        org.apache.poi.ss.usermodel.Hyperlink hyperlink = cell.getHyperlink();
        if (!hasUsableHyperlink(hyperlink)) {
          continue;
        }
        HyperlinkType hyperlinkType = hyperlink.getType();
        String address = cellAddress(cell);
        WorkbookAnalysis.AnalysisLocation.Cell location = cellLocation(address);
        findings.addAll(
            hyperlinkTargetFindings(
                location, hyperlinkType, hyperlink.getAddress(), workbookLocation));
      }
    }
    return List.copyOf(findings);
  }

  /** Returns hyperlink-health findings assuming the workbook has not yet been saved to disk. */
  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings() {
    return hyperlinkHealthFindings(new WorkbookLocation.UnsavedWorkbook());
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

  private ExcelCellSnapshot snapshotCellOrBlank(String address, int rowIndex, int columnIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      return blankSnapshot(address);
    }
    Cell cell = row.getCell(columnIndex);
    if (cell == null) {
      return blankSnapshot(address);
    }
    return snapshot(address, cell);
  }

  private void writeCellValue(int rowIndex, int columnIndex, ExcelCellValue value) {
    Cell cell = getOrCreateCell(rowIndex, columnIndex);

    switch (value) {
      case ExcelCellValue.BlankValue _ -> cell.setBlank();
      case ExcelCellValue.TextValue textValue -> cell.setCellValue(textValue.value());
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

  private void syncTableHeaders(ExcelRange changedRange) {
    ExcelTableHeaderSyncSupport.syncAffectedHeaders(xssfSheet(), changedRange);
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

  // Returns the cell at the given address, or null when the row or cell does not physically exist.
  private Cell optionalCell(String address) {
    CellReference cellReference = parseCellReference(address);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      return null;
    }
    return row.getCell(cellReference.getCol());
  }

  private Cell getOrCreateCell(int rowIndex, int columnIndex) {
    return getOrCreateRow(rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
  }

  XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private int nextAppendRowIndex() {
    int lastValueBearingRowIndex = -1;
    for (Row row : sheet) {
      if (rowHasValueBearingCell(row)) {
        lastValueBearingRowIndex = row.getRowNum();
      }
    }
    return lastValueBearingRowIndex + 1;
  }

  private boolean rowHasValueBearingCell(Row row) {
    for (Cell cell : row) {
      if (cell.getCellType() != CellType.BLANK) {
        return true;
      }
    }
    return false;
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
              ? formulaRuntime.displayValue(dataFormatter, cell)
              : dataFormatter.formatCellValue(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(name(), address, formulaExpression, exception);
    }
    ExcelCellStyleSnapshot style = styleRegistry.snapshot(cell);
    ExcelCellMetadataSnapshot metadata = metadata(cell);

    if (declaredType == CellType.FORMULA) {
      String formula = cell.getCellFormula();
      CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(name(), address, formula, exception);
      }

      CellType evalType = evaluatedCell != null ? evaluatedCell.getCellType() : CellType.BLANK;
      ExcelCellSnapshot evaluation =
          switch (evalType) {
            case STRING ->
                new ExcelCellSnapshot.TextSnapshot(
                    address,
                    "STRING",
                    displayValue,
                    style,
                    metadata,
                    evaluatedCell.getStringValue());
            case NUMERIC ->
                new ExcelCellSnapshot.NumberSnapshot(
                    address,
                    "NUMBER",
                    displayValue,
                    style,
                    metadata,
                    evaluatedCell.getNumberValue());
            case BOOLEAN ->
                new ExcelCellSnapshot.BooleanSnapshot(
                    address,
                    "BOOLEAN",
                    displayValue,
                    style,
                    metadata,
                    evaluatedCell.getBooleanValue());
            case ERROR ->
                new ExcelCellSnapshot.ErrorSnapshot(
                    address,
                    "ERROR",
                    displayValue,
                    style,
                    metadata,
                    FormulaError.forInt(evaluatedCell.getErrorValue()).getString());
            default ->
                new ExcelCellSnapshot.BlankSnapshot(
                    address, "BLANK", displayValue, style, metadata);
          };
      return new ExcelCellSnapshot.FormulaSnapshot(
          address, "FORMULA", displayValue, style, metadata, formula, evaluation);
    }

    return switch (declaredType) {
      case STRING ->
          new ExcelCellSnapshot.TextSnapshot(
              address, "STRING", displayValue, style, metadata, cell.getStringCellValue());
      case NUMERIC ->
          new ExcelCellSnapshot.NumberSnapshot(
              address, "NUMBER", displayValue, style, metadata, cell.getNumericCellValue());
      case BOOLEAN ->
          new ExcelCellSnapshot.BooleanSnapshot(
              address, "BOOLEAN", displayValue, style, metadata, cell.getBooleanCellValue());
      case ERROR ->
          new ExcelCellSnapshot.ErrorSnapshot(
              address,
              "ERROR",
              displayValue,
              style,
              metadata,
              FormulaError.forInt(cell.getErrorCellValue()).getString());
      case BLANK, _NONE, FORMULA ->
          new ExcelCellSnapshot.BlankSnapshot(address, "BLANK", displayValue, style, metadata);
    };
  }

  private void resetToDefaultStyle(Cell cell) {
    cell.setCellStyle(styleRegistry.defaultStyle());
  }

  private ExcelCellSnapshot blankSnapshot(String address) {
    return new ExcelCellSnapshot.BlankSnapshot(
        address, "BLANK", "", styleRegistry.defaultSnapshot(), ExcelCellMetadataSnapshot.empty());
  }

  private CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  private static void requireValidCellReference(String address, CellReference cellReference) {
    int row = cellReference.getRow();
    int col = cellReference.getCol();
    if (row < 0
        || col < 0
        || row > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || col > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new InvalidCellAddressException(
          address, new IllegalArgumentException("not a valid A1-style cell address: " + address));
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
          "widthCharacters must not exceed 255.0 (Excel column width limit): got "
              + widthCharacters);
    }
    int widthUnits = (int) Math.round(widthCharacters * 256.0d);
    if (widthUnits <= 0) {
      throw new IllegalArgumentException(
          "widthCharacters is too small to produce a visible Excel column width: got "
              + widthCharacters);
    }
    return widthUnits;
  }

  /** Converts row-height input to POI points after validating Excel storage constraints. */
  static float toRowHeightPoints(double heightPoints) {
    requireFinitePositive(heightPoints, "heightPoints");
    if (heightPoints > Short.MAX_VALUE / 20.0d) {
      throw new IllegalArgumentException(
          "heightPoints must not exceed 1638.35 (Excel storage limit: 32767 twips): got "
              + heightPoints);
    }
    if ((long) (heightPoints * 20.0d) <= 0L) {
      throw new IllegalArgumentException(
          "heightPoints is too small to produce a visible Excel row height: " + heightPoints);
    }
    return (float) heightPoints;
  }

  /**
   * Returns whether the cell should appear in preview output because it carries visible content or
   * metadata.
   */
  static boolean shouldPreview(Cell cell) {
    return cell != null
        && shouldPreview(
            cell.getCellType(),
            cell.getCellStyle().getIndex(),
            cell.getHyperlink() != null,
            cell.getCellComment() != null);
  }

  /** Returns whether the supplied cell facts would make a cell visible in preview output. */
  static boolean shouldPreview(
      CellType cellType, short styleIndex, boolean hasHyperlink, boolean hasComment) {
    return cellType != CellType.BLANK || styleIndex != 0 || hasHyperlink || hasComment;
  }

  private Comment newComment(int rowIndex, int columnIndex, ExcelComment comment) {
    ClientAnchor anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
    anchor.setRow1(rowIndex);
    anchor.setRow2(rowIndex + 3);
    anchor.setCol1(columnIndex);
    anchor.setCol2(columnIndex + 3);
    Comment poiComment = sheet.createDrawingPatriarch().createCellComment(anchor);
    poiComment.setAuthor(comment.author());
    poiComment.setVisible(comment.visible());
    poiComment.setString(
        sheet.getWorkbook().getCreationHelper().createRichTextString(comment.text()));
    return poiComment;
  }

  private ExcelCellMetadataSnapshot metadata(Cell cell) {
    return ExcelCellMetadataSnapshot.of(hyperlink(cell), comment(cell));
  }

  private List<WorkbookReadResult.CellHyperlink> allUsedHyperlinks() {
    List<WorkbookReadResult.CellHyperlink> hyperlinks = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        ExcelHyperlink hyperlink = hyperlink(cell);
        if (hyperlink != null) {
          hyperlinks.add(
              new WorkbookReadResult.CellHyperlink(
                  new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                  hyperlink));
        }
      }
    }
    return List.copyOf(hyperlinks);
  }

  private List<WorkbookReadResult.CellHyperlink> selectedHyperlinks(List<String> addresses) {
    List<WorkbookReadResult.CellHyperlink> hyperlinks = new ArrayList<>();
    for (String address : addresses) {
      Cell cell = cellOrNull(address);
      if (cell == null) {
        continue;
      }
      ExcelHyperlink hyperlink = hyperlink(cell);
      if (hyperlink != null) {
        hyperlinks.add(new WorkbookReadResult.CellHyperlink(address, hyperlink));
      }
    }
    return List.copyOf(hyperlinks);
  }

  private List<WorkbookReadResult.CellComment> allUsedComments() {
    List<WorkbookReadResult.CellComment> comments = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        ExcelComment comment = comment(cell);
        if (comment != null) {
          comments.add(
              new WorkbookReadResult.CellComment(
                  new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                  comment));
        }
      }
    }
    return List.copyOf(comments);
  }

  private List<WorkbookReadResult.CellComment> selectedComments(List<String> addresses) {
    List<WorkbookReadResult.CellComment> comments = new ArrayList<>();
    for (String address : addresses) {
      Cell cell = cellOrNull(address);
      if (cell == null) {
        continue;
      }
      ExcelComment comment = comment(cell);
      if (comment != null) {
        comments.add(new WorkbookReadResult.CellComment(address, comment));
      }
    }
    return List.copyOf(comments);
  }

  private Cell cellOrNull(String address) {
    CellReference reference = parseCellReference(address);
    Row row = sheet.getRow(reference.getRow());
    return row == null ? null : row.getCell(reference.getCol());
  }

  /**
   * Returns the normalized workbook-core hyperlink for the cell, or null when no usable hyperlink
   * exists.
   */
  static ExcelHyperlink hyperlink(Cell cell) {
    return cell == null ? null : hyperlink(cell.getHyperlink());
  }

  /**
   * Returns the normalized workbook-core hyperlink for the POI hyperlink, or null when unusable.
   */
  static ExcelHyperlink hyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    if (hyperlink == null || hyperlink.getType() == null) {
      return null;
    }
    return hyperlink(hyperlink.getType(), hyperlink.getAddress());
  }

  /**
   * Returns the normalized workbook-core hyperlink for the supplied type and target, or null when
   * unusable.
   */
  static ExcelHyperlink hyperlink(HyperlinkType hyperlinkType, String target) {
    if (hyperlinkType == null || target == null || target.isBlank()) {
      return null;
    }
    return switch (hyperlinkType) {
      case URL ->
          ExcelHyperlinkValidation.isValidUrlTarget(target) ? new ExcelHyperlink.Url(target) : null;
      case EMAIL ->
          ExcelHyperlinkValidation.isValidEmailTarget(target)
              ? new ExcelHyperlink.Email(target)
              : null;
      case FILE ->
          ExcelHyperlinkValidation.isValidFileTarget(target)
              ? new ExcelHyperlink.File(target)
              : null;
      case DOCUMENT -> new ExcelHyperlink.Document(target);
      case NONE -> null;
    };
  }

  /**
   * Returns the normalized workbook-core comment for the cell, or null when the POI comment is
   * incomplete.
   */
  static ExcelComment comment(Cell cell) {
    return cell == null ? null : comment(cell.getCellComment());
  }

  /** Returns the normalized workbook-core comment for the POI comment, or null when incomplete. */
  static ExcelComment comment(Comment comment) {
    if (comment == null || comment.getString() == null) {
      return null;
    }
    return comment(comment.getString().getString(), comment.getAuthor(), comment.isVisible());
  }

  /** Returns the normalized workbook-core comment for the supplied fields, or null when invalid. */
  static ExcelComment comment(String text, String author, boolean visible) {
    if (text == null || text.isBlank() || author == null || author.isBlank()) {
      return null;
    }
    return new ExcelComment(text, author, visible);
  }

  /** Converts the workbook-core hyperlink type into the matching Apache POI hyperlink type. */
  static HyperlinkType toPoi(ExcelHyperlinkType hyperlinkType) {
    return switch (hyperlinkType) {
      case URL -> HyperlinkType.URL;
      case EMAIL -> HyperlinkType.EMAIL;
      case FILE -> HyperlinkType.FILE;
      case DOCUMENT -> HyperlinkType.DOCUMENT;
    };
  }

  /** Converts the workbook-core hyperlink target into the exact Apache POI address string. */
  static String toPoiTarget(ExcelHyperlink hyperlink) {
    return switch (hyperlink) {
      case ExcelHyperlink.Url url -> url.target();
      case ExcelHyperlink.Email email -> "mailto:" + email.target();
      case ExcelHyperlink.File file -> ExcelFileHyperlinkTargets.toPoiAddress(file.path());
      case ExcelHyperlink.Document document -> document.target();
    };
  }

  /**
   * Returns analysis findings for one URL or EMAIL hyperlink target.
   *
   * <p>Package-private so tests can cover malformed external-target branches that Apache POI
   * rejects before they become workbook hyperlinks.
   */
  static List<WorkbookAnalysis.AnalysisFinding> externalHyperlinkFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      String expectedShape,
      boolean valid) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(expectedShape, "expectedShape must not be null");
    return valid
        ? List.of()
        : List.of(
            malformedHyperlinkFinding(
                location,
                "target does not match the expected " + expectedShape + " shape",
                List.of(target)));
  }

  /**
   * Returns analysis findings for one hyperlink target on this sheet.
   *
   * <p>Package-private so tests can cover blank and type-specific validation branches that Apache
   * POI does not always allow to be expressed as persisted workbook hyperlinks.
   */
  List<WorkbookAnalysis.AnalysisFinding> hyperlinkTargetFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      HyperlinkType hyperlinkType,
      String target,
      WorkbookLocation workbookLocation) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(hyperlinkType, "hyperlinkType must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    if (hasMissingHyperlinkTarget(target)) {
      return List.of(
          malformedHyperlinkFinding(
              location, "Hyperlink target is blank or missing.", List.of(hyperlinkType.name())));
    }
    return switch (hyperlinkType) {
      case URL ->
          externalHyperlinkFindings(
              location,
              target,
              ExcelHyperlinkType.URL.name(),
              ExcelHyperlinkValidation.isValidUrlTarget(target));
      case EMAIL ->
          externalHyperlinkFindings(
              location,
              target,
              ExcelHyperlinkType.EMAIL.name(),
              ExcelHyperlinkValidation.isValidEmailTarget(target));
      case FILE -> fileHyperlinkFindings(location, target, workbookLocation);
      case DOCUMENT -> {
        List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
        validateDocumentHyperlinkTarget(location, target, findings);
        yield List.copyOf(findings);
      }
      case NONE -> List.of();
    };
  }

  void validateDocumentHyperlinkTarget(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      List<WorkbookAnalysis.AnalysisFinding> findings) {
    int bangIndex = target.indexOf('!');
    if (bangIndex <= 0 || bangIndex >= target.length() - 1) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Invalid document hyperlink target",
              "Document hyperlink target must include a sheet and cell or range reference.",
              location,
              List.of(target)));
      return;
    }

    String targetSheetName = unquoteSheetName(target.substring(0, bangIndex));
    String targetAddress = target.substring(bangIndex + 1);
    if (sheet.getWorkbook().getSheet(targetSheetName) == null) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Document hyperlink targets a missing sheet",
              "Document hyperlink target sheet does not exist: " + targetSheetName,
              location,
              List.of(target)));
      return;
    }

    try {
      if (targetAddress.contains(":")) {
        ExcelRange.parse(targetAddress);
      } else {
        new CellReference(targetAddress);
      }
    } catch (IllegalArgumentException exception) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Invalid document hyperlink target",
              "Document hyperlink target is not a valid cell or range reference.",
              location,
              List.of(target)));
    }
  }

  private WorkbookAnalysis.AnalysisLocation.Cell cellLocation(String address) {
    return new WorkbookAnalysis.AnalysisLocation.Cell(name(), address);
  }

  private String cellAddress(Cell cell) {
    return cell.getAddress().formatAsString();
  }

  private static WorkbookAnalysis.AnalysisFinding malformedHyperlinkFinding(
      WorkbookAnalysis.AnalysisLocation.Cell location, String message, List<String> evidence) {
    return new WorkbookAnalysis.AnalysisFinding(
        WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
        WorkbookAnalysis.AnalysisSeverity.ERROR,
        "Malformed hyperlink target",
        "Hyperlink target is malformed: " + message,
        location,
        evidence);
  }

  /**
   * Returns analysis findings for one FILE hyperlink target in one workbook location context.
   *
   * <p>Package-private so tests can cover malformed and unresolved resolver branches that Apache
   * POI cannot always materialize as persisted hyperlink records.
   */
  static List<WorkbookAnalysis.AnalysisFinding> fileHyperlinkFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      WorkbookLocation workbookLocation) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    return switch (ExcelFileHyperlinkTargets.resolve(target, workbookLocation)) {
      case FileHyperlinkResolution.MalformedPath malformedPath ->
          List.of(
              malformedHyperlinkFinding(
                  location, malformedPath.reason(), List.of(malformedPath.path())));
      case FileHyperlinkResolution.UnresolvedRelativePath unresolvedRelativePath ->
          List.of(
              new WorkbookAnalysis.AnalysisFinding(
                  WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
                  WorkbookAnalysis.AnalysisSeverity.WARNING,
                  "Relative file hyperlink cannot be resolved yet",
                  "Relative file hyperlink targets cannot be validated until the workbook has a"
                      + " filesystem location.",
                  location,
                  List.of(unresolvedRelativePath.path())));
      case FileHyperlinkResolution.ResolvedPath resolvedPath ->
          Files.notExists(resolvedPath.resolvedPath())
              ? List.of(
                  new WorkbookAnalysis.AnalysisFinding(
                      WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET,
                      WorkbookAnalysis.AnalysisSeverity.ERROR,
                      "File hyperlink targets a missing path",
                      "Resolved file hyperlink path does not exist: " + resolvedPath.resolvedPath(),
                      location,
                      List.of(resolvedPath.path(), resolvedPath.resolvedPath().toString())))
              : List.of();
    };
  }

  String exceptionMessage(Exception exception) {
    if (exception.getMessage() == null) {
      return exception.getClass().getSimpleName();
    }
    return exception.getMessage();
  }

  static boolean hasUsableHyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return hyperlink != null && hasUsableHyperlinkType(hyperlink.getType());
  }

  /** Returns whether the supplied POI hyperlink type represents a usable hyperlink payload. */
  static boolean hasUsableHyperlinkType(HyperlinkType hyperlinkType) {
    return hyperlinkType != null && hyperlinkType != HyperlinkType.NONE;
  }

  static boolean hasMissingHyperlinkTarget(String target) {
    return target == null || target.isBlank();
  }

  static String unquoteSheetName(String sheetName) {
    if (!sheetName.startsWith("'")) {
      return sheetName;
    }
    if (!sheetName.endsWith("'")) {
      return sheetName;
    }
    if (sheetName.length() < 2) {
      return sheetName;
    }
    return sheetName.substring(1, sheetName.length() - 1).replace("''", "'");
  }

  static boolean containsExternalWorkbookReference(String formula) {
    return formula.contains("[") && formula.contains("]");
  }

  static List<String> volatileFunctions(String formula) {
    String upper = formula.toUpperCase(Locale.ROOT);
    List<String> functions = new ArrayList<>();
    for (String candidate :
        List.of("NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO")) {
      if (upper.contains(candidate + "(")) {
        functions.add(candidate);
      }
    }
    return List.copyOf(functions);
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
