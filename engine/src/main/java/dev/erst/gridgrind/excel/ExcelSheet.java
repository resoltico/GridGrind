package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import java.util.Objects;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * High-level sheet wrapper for typed reads, writes, and previews.
 *
 * <p>This facade intentionally owns the public sheet API while delegating behavior into narrower
 * support classes; the public-method count reflects that single boundary.
 */
@SuppressWarnings("PMD.ExcessivePublicCount")
public final class ExcelSheet {
  private final Sheet sheet;
  private final ExcelSheetDrawingSupport drawingSupport;
  private final ExcelSheetAnnotationSupport annotationSupport;
  private final ExcelSheetMetadataSupport metadataSupport;
  private final ExcelSheetStructureSupport structureSupport;
  private final ExcelSheetAnalysisSupport analysisSupport;
  private final ExcelSheetCellMutationSupport mutationSupport;
  private final ExcelSheetCellReadSupport readSupport;

  ExcelSheet(Sheet sheet, WorkbookStyleRegistry styleRegistry, ExcelFormulaRuntime formulaRuntime) {
    this.sheet = sheet;
    DataFormatter dataFormatter = new DataFormatter();
    ExcelDataValidationController dataValidationController = new ExcelDataValidationController();
    ExcelConditionalFormattingController conditionalFormattingController =
        new ExcelConditionalFormattingController();
    ExcelAutofilterController autofilterController = new ExcelAutofilterController();
    ExcelPrintLayoutController printLayoutController = new ExcelPrintLayoutController();
    ExcelSheetPresentationController sheetPresentationController =
        new ExcelSheetPresentationController();
    ExcelRowColumnStructureController rowColumnStructureController =
        new ExcelRowColumnStructureController();
    ExcelDrawingController drawingController = new ExcelDrawingController();
    this.drawingSupport = new ExcelSheetDrawingSupport(sheet, drawingController, formulaRuntime);
    this.annotationSupport = new ExcelSheetAnnotationSupport(sheet, drawingController);
    this.metadataSupport =
        new ExcelSheetMetadataSupport(
            sheet, dataValidationController, conditionalFormattingController, autofilterController);
    this.structureSupport =
        new ExcelSheetStructureSupport(
            sheet,
            formulaRuntime,
            dataFormatter,
            printLayoutController,
            sheetPresentationController,
            rowColumnStructureController);
    this.analysisSupport = new ExcelSheetAnalysisSupport(sheet, formulaRuntime);
    this.mutationSupport =
        new ExcelSheetCellMutationSupport(sheet, styleRegistry, formulaRuntime, drawingController);
    this.readSupport =
        new ExcelSheetCellReadSupport(
            sheet, styleRegistry, formulaRuntime, dataFormatter, annotationSupport);
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
    mutationSupport.setCell(address, value);
    return this;
  }

  /** Writes a rectangular matrix of values to an A1-style range such as {@code A1:C3}. */
  public ExcelSheet setRange(String range, List<List<ExcelCellValue>> rows) {
    mutationSupport.setRange(range, rows);
    return this;
  }

  /** Clears both contents and formatting in a rectangular A1-style range. */
  public ExcelSheet clearRange(String range) {
    mutationSupport.clearRange(range);
    return this;
  }

  /** Creates or replaces one dedicated array-formula group over a rectangular range. */
  public ExcelSheet setArrayFormula(String range, ExcelArrayFormulaDefinition formula) {
    mutationSupport.setArrayFormula(range, formula);
    return this;
  }

  /** Removes the array-formula group containing one addressed cell. */
  public ExcelSheet clearArrayFormula(String address) {
    mutationSupport.clearArrayFormula(address);
    return this;
  }

  /** Replaces the hyperlink attached to one cell, creating the cell if necessary. */
  public ExcelSheet setHyperlink(String address, ExcelHyperlink hyperlink) {
    annotationSupport.setHyperlink(address, hyperlink);
    return this;
  }

  /** Removes any hyperlink attached to one cell; no-op when the cell does not physically exist. */
  public ExcelSheet clearHyperlink(String address) {
    annotationSupport.clearHyperlink(address);
    return this;
  }

  /** Replaces the plain-text comment attached to one cell, creating the cell if necessary. */
  public ExcelSheet setComment(String address, ExcelComment comment) {
    annotationSupport.setComment(address, comment);
    return this;
  }

  /** Removes any comment attached to one cell; no-op when the cell does not physically exist. */
  public ExcelSheet clearComment(String address) {
    annotationSupport.clearComment(address);
    return this;
  }

  /** Creates or replaces one picture-backed drawing object on this sheet. */
  public ExcelSheet setPicture(ExcelPictureDefinition definition) {
    return drawingSupport.setPicture(definition, this);
  }

  /** Creates or replaces one signature-line drawing object on this sheet. */
  public ExcelSheet setSignatureLine(ExcelSignatureLineDefinition definition) {
    return drawingSupport.setSignatureLine(definition, this);
  }

  /** Creates or mutates one supported simple chart on this sheet. */
  public ExcelSheet setChart(ExcelChartDefinition definition) {
    return drawingSupport.setChart(definition, this);
  }

  /** Creates or replaces one simple-shape or connector drawing object on this sheet. */
  public ExcelSheet setShape(ExcelShapeDefinition definition) {
    return drawingSupport.setShape(definition, this);
  }

  /** Creates or replaces one embedded-object drawing object on this sheet. */
  public ExcelSheet setEmbeddedObject(ExcelEmbeddedObjectDefinition definition) {
    return drawingSupport.setEmbeddedObject(definition, this);
  }

  /** Moves one existing drawing object by replacing its anchor authoritatively. */
  public ExcelSheet setDrawingObjectAnchor(String objectName, ExcelDrawingAnchor.TwoCell anchor) {
    return drawingSupport.setDrawingObjectAnchor(objectName, anchor, this);
  }

  /** Deletes one existing drawing object by sheet-local name. */
  public ExcelSheet deleteDrawingObject(String objectName) {
    return drawingSupport.deleteDrawingObject(objectName, this);
  }

  /** Applies a style patch to every cell in a rectangular A1-style range. */
  public ExcelSheet applyStyle(String range, ExcelCellStyle style) {
    mutationSupport.applyStyle(range, style);
    return this;
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  public ExcelSheet setDataValidation(String range, ExcelDataValidationDefinition validation) {
    return metadataSupport.setDataValidation(range, validation, this);
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  public ExcelSheet clearDataValidations(ExcelRangeSelection selection) {
    return metadataSupport.clearDataValidations(selection, this);
  }

  /** Creates or replaces one logical conditional-formatting block on this sheet. */
  public ExcelSheet setConditionalFormatting(ExcelConditionalFormattingBlockDefinition block) {
    return metadataSupport.setConditionalFormatting(block, this);
  }

  /** Removes conditional-formatting blocks on this sheet matching the provided selection. */
  public ExcelSheet clearConditionalFormatting(ExcelRangeSelection selection) {
    return metadataSupport.clearConditionalFormatting(selection, this);
  }

  /** Creates or replaces one sheet-level autofilter range. */
  public ExcelSheet setAutofilter(String range) {
    return metadataSupport.setAutofilter(range, this);
  }

  /** Creates or replaces one sheet-level autofilter range plus authored criteria and sort state. */
  public ExcelSheet setAutofilter(
      String range,
      List<ExcelAutofilterFilterColumn> criteria,
      ExcelAutofilterSortState sortState) {
    return metadataSupport.setAutofilter(range, criteria, sortState, this);
  }

  /** Clears the sheet-level autofilter range on this sheet. */
  public ExcelSheet clearAutofilter() {
    return metadataSupport.clearAutofilter(this);
  }

  /** Appends a new row using the next available row index. */
  public ExcelSheet appendRow(ExcelCellValue... values) {
    mutationSupport.appendRow(values);
    return this;
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  public ExcelSheet mergeCells(String range) {
    return structureSupport.mergeCells(range, this);
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  public ExcelSheet unmergeCells(String range) {
    return structureSupport.unmergeCells(range, this);
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  public ExcelSheet setColumnWidth(
      int firstColumnIndex, int lastColumnIndex, double widthCharacters) {
    return structureSupport.setColumnWidth(
        firstColumnIndex, lastColumnIndex, widthCharacters, this);
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  public ExcelSheet setRowHeight(int firstRowIndex, int lastRowIndex, double heightPoints) {
    return structureSupport.setRowHeight(firstRowIndex, lastRowIndex, heightPoints, this);
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  public ExcelSheet insertRows(int rowIndex, int rowCount) {
    return structureSupport.insertRows(rowIndex, rowCount, this);
  }

  /** Deletes the requested inclusive zero-based row band. */
  public ExcelSheet deleteRows(ExcelRowSpan rows) {
    return structureSupport.deleteRows(rows, this);
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  public ExcelSheet shiftRows(ExcelRowSpan rows, int delta) {
    return structureSupport.shiftRows(rows, delta, this);
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  public ExcelSheet insertColumns(int columnIndex, int columnCount) {
    return structureSupport.insertColumns(columnIndex, columnCount, this);
  }

  /** Deletes the requested inclusive zero-based column band. */
  public ExcelSheet deleteColumns(ExcelColumnSpan columns) {
    return structureSupport.deleteColumns(columns, this);
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  public ExcelSheet shiftColumns(ExcelColumnSpan columns, int delta) {
    return structureSupport.shiftColumns(columns, delta, this);
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  public ExcelSheet setRowVisibility(ExcelRowSpan rows, boolean hidden) {
    return structureSupport.setRowVisibility(rows, hidden, this);
  }

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  public ExcelSheet setColumnVisibility(ExcelColumnSpan columns, boolean hidden) {
    return structureSupport.setColumnVisibility(columns, hidden, this);
  }

  /** Applies one outline group to the requested inclusive zero-based row band. */
  public ExcelSheet groupRows(ExcelRowSpan rows, boolean collapsed) {
    return structureSupport.groupRows(rows, collapsed, this);
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  public ExcelSheet ungroupRows(ExcelRowSpan rows) {
    return structureSupport.ungroupRows(rows, this);
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  public ExcelSheet groupColumns(ExcelColumnSpan columns, boolean collapsed) {
    return structureSupport.groupColumns(columns, collapsed, this);
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  public ExcelSheet ungroupColumns(ExcelColumnSpan columns) {
    return structureSupport.ungroupColumns(columns, this);
  }

  /** Applies one explicit pane state to this sheet. */
  public ExcelSheet setPane(ExcelSheetPane pane) {
    return structureSupport.setPane(pane, this);
  }

  /** Applies one explicit zoom percentage to this sheet. */
  public ExcelSheet setZoom(int zoomPercent) {
    return structureSupport.setZoom(zoomPercent, this);
  }

  /** Applies authoritative sheet-presentation state such as display flags and defaults. */
  public ExcelSheet setPresentation(ExcelSheetPresentation presentation) {
    return structureSupport.setPresentation(presentation, this);
  }

  /** Applies the provided print layout as the authoritative supported print state. */
  public ExcelSheet setPrintLayout(ExcelPrintLayout printLayout) {
    return structureSupport.setPrintLayout(printLayout, this);
  }

  /** Clears the supported print layout state from this sheet. */
  public ExcelSheet clearPrintLayout() {
    return structureSupport.clearPrintLayout(this);
  }

  /** Returns the number of physically stored rows in the sheet. */
  public int physicalRowCount() {
    return structureSupport.physicalRowCount();
  }

  /** Returns the 0-based index of the last row, or -1 if no rows exist. */
  public int lastRowIndex() {
    return structureSupport.lastRowIndex();
  }

  /**
   * Returns the 0-based index of the widest column across all rows, or -1 if the sheet is empty.
   */
  public int lastColumnIndex() {
    return structureSupport.lastColumnIndex();
  }

  /** Auto-sizes all populated columns on this sheet to fit their content. */
  public ExcelSheet autoSizeColumns() {
    return structureSupport.autoSizeColumns(this, name());
  }

  /** Reads a string cell by A1-style address. */
  public String text(String address) {
    return readSupport.text(address);
  }

  /** Reads a numeric cell by A1-style address, evaluating formulas when needed. */
  public double number(String address) {
    return readSupport.number(address);
  }

  /** Reads a boolean cell by A1-style address, evaluating formulas when needed. */
  public boolean bool(String address) {
    return readSupport.bool(address);
  }

  /** Reads the raw formula expression stored in a formula cell. */
  public String formula(String address) {
    return readSupport.formula(address);
  }

  /** Captures a formatted, typed snapshot of a single cell, returning blank for unwritten cells. */
  public ExcelCellSnapshot snapshotCell(String address) {
    return readSupport.snapshotCell(address);
  }

  /** Captures exact snapshots for the provided ordered A1 addresses. */
  public List<ExcelCellSnapshot> snapshotCells(List<String> addresses) {
    return readSupport.snapshotCells(addresses);
  }

  /** Returns a compact preview of the top-left portion of the sheet. */
  public List<ExcelPreviewRow> preview(int maxRows, int maxColumns) {
    return readSupport.preview(maxRows, maxColumns, lastRowIndex());
  }

  /** Returns an exact rectangular window of cell snapshots anchored at one top-left address. */
  public WorkbookSheetResult.Window window(String topLeftAddress, int rowCount, int columnCount) {
    return readSupport.window(topLeftAddress, rowCount, columnCount);
  }

  /** Returns every merged region currently defined on the sheet. */
  public List<WorkbookSheetResult.MergedRegion> mergedRegions() {
    return structureSupport.mergedRegions();
  }

  /** Returns hyperlink metadata for the selected cells on this sheet. */
  public List<WorkbookSheetResult.CellHyperlink> hyperlinks(ExcelCellSelection selection) {
    return annotationSupport.hyperlinks(selection);
  }

  /** Returns comment metadata for the selected cells on this sheet. */
  public List<WorkbookSheetResult.CellComment> comments(ExcelCellSelection selection) {
    return annotationSupport.comments(selection);
  }

  /** Returns layout metadata such as panes, zoom, and visible sizing. */
  public WorkbookSheetResult.SheetLayout layout() {
    return structureSupport.layout(name());
  }

  /** Returns supported print-layout metadata for this sheet. */
  public ExcelPrintLayout printLayout() {
    return structureSupport.printLayout();
  }

  /** Returns the full factual print-layout snapshot currently stored for this sheet. */
  public ExcelPrintLayoutSnapshot printLayoutSnapshot() {
    return structureSupport.printLayoutSnapshot();
  }

  /** Returns data-validation metadata for the selected ranges on this sheet. */
  public List<ExcelDataValidationSnapshot> dataValidations(ExcelRangeSelection selection) {
    return metadataSupport.dataValidations(selection);
  }

  /** Returns factual conditional-formatting blocks for the selected ranges on this sheet. */
  public List<ExcelConditionalFormattingBlockSnapshot> conditionalFormatting(
      ExcelRangeSelection selection) {
    return metadataSupport.conditionalFormatting(selection);
  }

  /** Returns factual drawing-object metadata for this sheet. */
  public List<ExcelDrawingObjectSnapshot> drawingObjects() {
    return drawingSupport.drawingObjects();
  }

  /** Returns factual chart metadata for this sheet. */
  public List<ExcelChartSnapshot> charts() {
    return drawingSupport.charts();
  }

  /** Returns the extracted binary payload for one existing drawing object on this sheet. */
  public ExcelDrawingObjectPayload drawingObjectPayload(String objectName) {
    return drawingSupport.drawingObjectPayload(objectName);
  }

  /** Returns every formula cell currently present on the sheet. */
  public List<ExcelCellSnapshot.FormulaSnapshot> formulaCells() {
    return readSupport.formulaCells();
  }

  /** Returns factual array-formula groups on this sheet. */
  public List<ExcelArrayFormulaSnapshot> arrayFormulas() {
    return readSupport.arrayFormulas();
  }

  /** Returns the number of formula cells currently present on the sheet. */
  int formulaCellCount() {
    return analysisSupport.formulaCellCount();
  }

  /** Returns the number of conditional-formatting blocks currently present on the sheet. */
  int conditionalFormattingBlockCount() {
    return metadataSupport.conditionalFormattingBlockCount();
  }

  /** Returns derived findings for formula health on this sheet. */
  List<WorkbookAnalysis.AnalysisFinding> formulaHealthFindings() {
    return analysisSupport.formulaHealthFindings(name());
  }

  /** Returns derived conditional-formatting health findings on this sheet. */
  List<WorkbookAnalysis.AnalysisFinding> conditionalFormattingHealthFindings() {
    return metadataSupport.conditionalFormattingHealthFindings(name());
  }

  /** Returns the number of raw hyperlinks currently present on the sheet. */
  int hyperlinkCount() {
    return analysisSupport.hyperlinkCount();
  }

  /**
   * Returns derived findings for hyperlink health on this sheet using the workbook path context.
   */
  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings(
      WorkbookLocation workbookLocation) {
    return analysisSupport.hyperlinkHealthFindings(workbookLocation);
  }

  /** Returns hyperlink-health findings assuming the workbook has not yet been saved to disk. */
  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings() {
    return analysisSupport.hyperlinkHealthFindings();
  }

  XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
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

  /**
   * Returns the normalized workbook-core hyperlink for the cell, or null when no usable hyperlink
   * exists.
   */
  static ExcelHyperlink hyperlink(Cell cell) {
    return ExcelSheetAnnotationSupport.hyperlink(cell);
  }

  /**
   * Returns the normalized workbook-core hyperlink for the POI hyperlink, or null when unusable.
   */
  static ExcelHyperlink hyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return ExcelSheetAnnotationSupport.hyperlink(hyperlink);
  }

  /**
   * Returns the normalized workbook-core hyperlink for the supplied type and target, or null when
   * unusable.
   */
  static ExcelHyperlink hyperlink(HyperlinkType hyperlinkType, String target) {
    return ExcelSheetAnnotationSupport.hyperlink(hyperlinkType, target);
  }

  /**
   * Returns the normalized workbook-core comment for the cell, or null when the POI comment is
   * incomplete.
   */
  static ExcelComment comment(Cell cell) {
    return ExcelSheetAnnotationSupport.comment(cell);
  }

  /** Returns the normalized workbook-core comment for the POI comment, or null when incomplete. */
  static ExcelComment comment(Comment comment) {
    return ExcelSheetAnnotationSupport.comment(comment);
  }

  /** Returns the normalized workbook-core comment for the supplied fields, or null when invalid. */
  static ExcelComment comment(String text, String author, boolean visible) {
    return ExcelSheetAnnotationSupport.comment(text, author, visible);
  }

  /**
   * Returns the full factual workbook-core comment snapshot for the cell, or null when incomplete.
   */
  static ExcelCommentSnapshot commentSnapshot(Cell cell) {
    return ExcelSheetAnnotationSupport.commentSnapshot(cell);
  }

  /**
   * Returns the full factual workbook-core comment snapshot for the POI comment, or null when
   * incomplete.
   */
  static ExcelCommentSnapshot commentSnapshot(Comment comment) {
    return ExcelSheetAnnotationSupport.commentSnapshot(comment);
  }

  /** Converts the workbook-core hyperlink type into the matching Apache POI hyperlink type. */
  static HyperlinkType toPoi(ExcelHyperlinkType hyperlinkType) {
    return ExcelSheetAnnotationSupport.toPoi(hyperlinkType);
  }

  /** Converts the workbook-core hyperlink target into the exact Apache POI address string. */
  static String toPoiTarget(ExcelHyperlink hyperlink) {
    return ExcelSheetAnnotationSupport.toPoiTarget(hyperlink);
  }

  static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  static int findMergedRegionIndex(Sheet sheet, ExcelRange excelRange) {
    return ExcelSheetStructureSupport.findMergedRegionIndex(sheet, excelRange);
  }

  static void requireNoMergedRegionOverlap(Sheet sheet, ExcelRange excelRange) {
    ExcelSheetStructureSupport.requireNoMergedRegionOverlap(sheet, excelRange);
  }

  static boolean matches(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return ExcelSheetStructureSupport.matches(rangeAddress, excelRange);
  }

  static boolean intersects(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return ExcelSheetStructureSupport.intersects(rangeAddress, excelRange);
  }

  static int toColumnWidthUnits(double widthCharacters) {
    return ExcelSheetStructureSupport.toColumnWidthUnits(widthCharacters);
  }

  static float toRowHeightPoints(double heightPoints) {
    return ExcelSheetStructureSupport.toRowHeightPoints(heightPoints);
  }

  static List<WorkbookAnalysis.AnalysisFinding> externalHyperlinkFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      String expectedShape,
      boolean valid) {
    return ExcelSheetAnalysisSupport.externalHyperlinkFindings(
        location, target, expectedShape, valid);
  }

  static List<WorkbookAnalysis.AnalysisFinding> fileHyperlinkFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      WorkbookLocation workbookLocation) {
    return ExcelSheetAnalysisSupport.fileHyperlinkFindings(location, target, workbookLocation);
  }

  static boolean hasUsableHyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return ExcelSheetAnalysisSupport.hasUsableHyperlink(hyperlink);
  }

  static boolean hasUsableHyperlinkType(HyperlinkType hyperlinkType) {
    return ExcelSheetAnalysisSupport.hasUsableHyperlinkType(hyperlinkType);
  }

  static boolean hasMissingHyperlinkTarget(String target) {
    return ExcelSheetAnalysisSupport.hasMissingHyperlinkTarget(target);
  }

  static String unquoteSheetName(String sheetName) {
    return ExcelSheetAnalysisSupport.unquoteSheetName(sheetName);
  }

  static boolean containsExternalWorkbookReference(String formula) {
    return ExcelSheetAnalysisSupport.containsExternalWorkbookReference(formula);
  }

  static List<String> volatileFunctions(String formula) {
    return ExcelSheetAnalysisSupport.volatileFunctions(formula);
  }

  String exceptionMessage(Exception exception) {
    return ExcelSheetAnalysisSupport.exceptionMessage(exception);
  }

  List<WorkbookAnalysis.AnalysisFinding> hyperlinkTargetFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      HyperlinkType hyperlinkType,
      String target,
      WorkbookLocation workbookLocation) {
    return analysisSupport.hyperlinkTargetFindings(
        location, hyperlinkType, target, workbookLocation);
  }

  void validateDocumentHyperlinkTarget(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      List<WorkbookAnalysis.AnalysisFinding> findings) {
    analysisSupport.validateDocumentHyperlinkTarget(location, target, findings);
  }
}
