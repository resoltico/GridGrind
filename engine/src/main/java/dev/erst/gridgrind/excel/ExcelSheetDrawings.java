package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import java.util.List;
import java.util.Objects;

/** Drawing, chart, picture, and embedded-object operations for one sheet. */
public final class ExcelSheetDrawings {
  private final ExcelSheet sheet;
  private final ExcelSheetDrawingSupport drawingSupport;

  ExcelSheetDrawings(ExcelSheet sheet, ExcelSheetDrawingSupport drawingSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.drawingSupport = Objects.requireNonNull(drawingSupport, "drawingSupport must not be null");
  }

  /** Creates or replaces one picture-backed drawing object on this sheet. */
  public ExcelSheetDrawings setPicture(ResolvedPicture definition) {
    drawingSupport.setPicture(definition, sheet);
    return this;
  }

  /** Creates or replaces one signature-line drawing object on this sheet. */
  public ExcelSheetDrawings setSignatureLine(ResolvedSignatureLine definition) {
    drawingSupport.setSignatureLine(definition, sheet);
    return this;
  }

  /** Creates or mutates one supported simple chart on this sheet. */
  public ExcelSheetDrawings setChart(ResolvedChart definition) {
    drawingSupport.setChart(definition, sheet);
    return this;
  }

  /** Creates or replaces one simple-shape or connector drawing object on this sheet. */
  public ExcelSheetDrawings setShape(ResolvedShape definition) {
    drawingSupport.setShape(definition, sheet);
    return this;
  }

  /** Creates or replaces one embedded-object drawing object on this sheet. */
  public ExcelSheetDrawings setEmbeddedObject(ResolvedEmbeddedObject definition) {
    drawingSupport.setEmbeddedObject(definition, sheet);
    return this;
  }

  /** Moves one existing drawing object by replacing its anchor authoritatively. */
  public ExcelSheetDrawings setDrawingObjectAnchor(
      String objectName, DrawingAnchorSpec.TwoCell anchor) {
    drawingSupport.setDrawingObjectAnchor(objectName, anchor, sheet);
    return this;
  }

  /** Deletes one existing drawing object by sheet-local name. */
  public ExcelSheetDrawings deleteDrawingObject(String objectName) {
    drawingSupport.deleteDrawingObject(objectName, sheet);
    return this;
  }

  /** Returns factual drawing-object metadata for this sheet. */
  public List<ExcelDrawingObjectSnapshot> drawingObjects() {
    return drawingSupport.drawingObjects();
  }

  /** Returns factual chart metadata for this sheet. */
  public List<ChartReport> charts() {
    return drawingSupport.charts();
  }

  /** Returns the extracted binary payload for one existing drawing object on this sheet. */
  public ExcelDrawingObjectPayload drawingObjectPayload(String objectName) {
    return drawingSupport.drawingObjectPayload(objectName);
  }
}
