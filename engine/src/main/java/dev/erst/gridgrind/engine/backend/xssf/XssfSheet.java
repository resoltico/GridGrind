package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import dev.erst.gridgrind.engine.cells.CellsAdapter;
import dev.erst.gridgrind.engine.comments.CommentAdapter;
import dev.erst.gridgrind.engine.conditionalformatting.ConditionalFormattingAdapter;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.layout.ColumnAdapter;
import dev.erst.gridgrind.engine.layout.LayoutAdapter;
import dev.erst.gridgrind.engine.layout.RowAdapter;
import dev.erst.gridgrind.engine.links.LinkAdapter;
import dev.erst.gridgrind.engine.print.PrintAdapter;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import dev.erst.gridgrind.engine.validation.DataValidationAdapter;
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
public final class XssfSheet {
  private final Sheet sheet;
  private final CellsAdapter cells;
  private final CommentAdapter comments;
  private final LinkAdapter links;
  private final SheetDrawings drawings;
  private final SheetMetadata metadata;
  private final LayoutAdapter layout;
  private final RowAdapter rows;
  private final ColumnAdapter columns;
  private final SheetDiagnostics diagnostics;

  XssfSheet(Sheet sheet, StyleAdapter styleRegistry, FormulaRuntime formulaRuntime) {
    this.sheet = sheet;
    DataFormatter dataFormatter = new DataFormatter();
    DataValidationAdapter dataValidationAdapter = new DataValidationAdapter();
    ConditionalFormattingAdapter conditionalFormattingAdapter = new ConditionalFormattingAdapter();
    AutofilterAdapter autofilterAdapter = new AutofilterAdapter();
    PrintAdapter printAdapter = new PrintAdapter();
    DrawingAdapter drawingController = new DrawingAdapter();
    SheetDrawings drawingSupport = new SheetDrawings(sheet, drawingController, formulaRuntime);
    CommentAdapter commentAdapter = new CommentAdapter((XSSFSheet) sheet);
    LinkAdapter linkAdapter = new LinkAdapter(sheet);
    SheetMetadata metadataSupport =
        new SheetMetadata(
            sheet, dataValidationAdapter, conditionalFormattingAdapter, autofilterAdapter);
    LayoutAdapter layoutAdapter = new LayoutAdapter(sheet, printAdapter);
    this.cells = new CellsAdapter(sheet, styleRegistry, formulaRuntime, dataFormatter);
    this.comments = commentAdapter;
    this.links = linkAdapter;
    this.drawings = drawingSupport;
    this.metadata = metadataSupport;
    this.layout = layoutAdapter;
    this.rows = new RowAdapter(sheet);
    this.columns = new ColumnAdapter(sheet, dataFormatter);
    this.diagnostics = new SheetDiagnostics(this, formulaRuntime, metadataSupport);
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  XssfSheet(Sheet sheet, StyleAdapter styleRegistry, FormulaEvaluator formulaEvaluator) {
    this(sheet, styleRegistry, FormulaRuntime.poi(formulaEvaluator));
  }

  /** Returns the sheet name as defined in the workbook. */
  public String name() {
    return sheet.getSheetName();
  }

  /** Returns the grouped cell read, write, preview, and array-formula surface. */
  public CellsAdapter cells() {
    return cells;
  }

  /** Returns comment operations bound to this sheet. */
  public CommentAdapter comments() {
    return comments;
  }

  /** Returns hyperlink operations bound to this sheet. */
  public LinkAdapter links() {
    return links;
  }

  /** Returns the grouped drawing, chart, and embedded-object surface. */
  public SheetDrawings drawings() {
    return drawings;
  }

  /** Returns the grouped validation, conditional-formatting, and autofilter surface. */
  public SheetMetadata metadata() {
    return metadata;
  }

  /** Returns the grouped layout, merge, pane, zoom, and print surface. */
  public LayoutAdapter layout() {
    return layout;
  }

  /** Returns the grouped row insertion, movement, visibility, and grouping surface. */
  public RowAdapter rows() {
    return rows;
  }

  /** Returns the grouped column insertion, movement, visibility, and sizing surface. */
  public ColumnAdapter columns() {
    return columns;
  }

  SheetDiagnostics diagnostics() {
    return diagnostics;
  }

  XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  String exceptionMessage(Exception exception) {
    return FormulaAdapter.exceptionMessage(exception);
  }

  java.util.List<AnalysisFindingReport> hyperlinkTargetFindings(
      AnalysisLocationReport.Cell location,
      HyperlinkType hyperlinkType,
      String target,
      WorkbookLocation workbookLocation) {
    return links.hyperlinkTargetFindings(location, hyperlinkType, target, workbookLocation);
  }

  void validateDocumentHyperlinkTarget(
      AnalysisLocationReport.Cell location,
      String target,
      java.util.List<AnalysisFindingReport> findings) {
    links.validateDocumentHyperlinkTarget(location, target, findings);
  }
}
