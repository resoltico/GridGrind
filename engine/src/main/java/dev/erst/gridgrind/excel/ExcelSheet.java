package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationController;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * High-level sheet wrapper for typed reads, writes, and previews.
 *
 * <p>This facade intentionally owns the public sheet API while delegating behavior into narrower
 * support classes; the public-method count reflects that single boundary.
 */
public final class ExcelSheet {
  private final Sheet sheet;
  private final ExcelSheetCells cells;
  private final ExcelSheetAnnotations annotations;
  private final ExcelSheetDrawings drawings;
  private final ExcelSheetMetadata metadata;
  private final ExcelSheetLayout layout;
  private final ExcelSheetRows rows;
  private final ExcelSheetColumns columns;
  private final ExcelSheetDiagnostics diagnostics;
  private final ExcelSheetAnalysisSupport analysisSupport;

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
    ExcelRowStructureController rowStructureController = new ExcelRowStructureController();
    ExcelColumnStructureController columnStructureController = new ExcelColumnStructureController();
    ExcelDrawingController drawingController = new ExcelDrawingController();
    ExcelSheetDrawingSupport drawingSupport =
        new ExcelSheetDrawingSupport(sheet, drawingController, formulaRuntime);
    ExcelSheetAnnotationSupport annotationSupport =
        new ExcelSheetAnnotationSupport(sheet, drawingController);
    ExcelSheetMetadataSupport metadataSupport =
        new ExcelSheetMetadataSupport(
            sheet, dataValidationController, conditionalFormattingController, autofilterController);
    ExcelSheetLayoutSupport layoutSupport =
        new ExcelSheetLayoutSupport(
            sheet,
            printLayoutController,
            sheetPresentationController,
            rowStructureController,
            columnStructureController);
    ExcelSheetRowSupport rowSupport = new ExcelSheetRowSupport(sheet, rowStructureController);
    ExcelSheetColumnSupport columnSupport =
        new ExcelSheetColumnSupport(
            sheet, formulaRuntime, dataFormatter, columnStructureController);
    this.analysisSupport = new ExcelSheetAnalysisSupport(sheet, formulaRuntime);
    ExcelSheetCellMutationSupport mutationSupport =
        new ExcelSheetCellMutationSupport(sheet, styleRegistry, formulaRuntime, drawingController);
    ExcelSheetCellReadSupport readSupport =
        new ExcelSheetCellReadSupport(
            sheet, styleRegistry, formulaRuntime, dataFormatter, annotationSupport);
    this.cells = new ExcelSheetCells(this, mutationSupport, readSupport);
    this.annotations = new ExcelSheetAnnotations(this, annotationSupport);
    this.drawings = new ExcelSheetDrawings(this, drawingSupport);
    this.metadata = new ExcelSheetMetadata(this, metadataSupport);
    this.layout = new ExcelSheetLayout(this, layoutSupport);
    this.rows = new ExcelSheetRows(this, rowSupport);
    this.columns = new ExcelSheetColumns(this, columnSupport);
    this.diagnostics = new ExcelSheetDiagnostics(this, analysisSupport, metadataSupport);
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  ExcelSheet(Sheet sheet, WorkbookStyleRegistry styleRegistry, FormulaEvaluator formulaEvaluator) {
    this(sheet, styleRegistry, ExcelFormulaRuntime.poi(formulaEvaluator));
  }

  /** Returns the sheet name as defined in the workbook. */
  public String name() {
    return sheet.getSheetName();
  }

  /** Returns the grouped cell read, write, preview, and array-formula surface. */
  public ExcelSheetCells cells() {
    return cells;
  }

  /** Returns the grouped hyperlink and comment surface. */
  public ExcelSheetAnnotations annotations() {
    return annotations;
  }

  /** Returns the grouped drawing, chart, and embedded-object surface. */
  public ExcelSheetDrawings drawings() {
    return drawings;
  }

  /** Returns the grouped validation, conditional-formatting, and autofilter surface. */
  public ExcelSheetMetadata metadata() {
    return metadata;
  }

  /** Returns the grouped layout, merge, pane, zoom, and print surface. */
  public ExcelSheetLayout layout() {
    return layout;
  }

  /** Returns the grouped row insertion, movement, visibility, and grouping surface. */
  public ExcelSheetRows rows() {
    return rows;
  }

  /** Returns the grouped column insertion, movement, visibility, and sizing surface. */
  public ExcelSheetColumns columns() {
    return columns;
  }

  ExcelSheetDiagnostics diagnostics() {
    return diagnostics;
  }

  XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  static void requireNonBlank(String value, String fieldName) {
    ExcelArgumentSupport.requireNonBlank(value, fieldName);
  }

  String exceptionMessage(Exception exception) {
    return ExcelSheetAnalysisSupport.exceptionMessage(exception);
  }

  java.util.List<WorkbookAnalysis.AnalysisFinding> hyperlinkTargetFindings(
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
      java.util.List<WorkbookAnalysis.AnalysisFinding> findings) {
    analysisSupport.validateDocumentHyperlinkTarget(location, target, findings);
  }
}
