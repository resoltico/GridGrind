package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;

/** Drawing-domain mutations and readback for one sheet wrapper. */
public final class SheetDrawings {
  private final Sheet sheet;
  private final DrawingAdapter drawingController;
  private final @Nullable FormulaRuntime formulaRuntime;

  /** Binds drawing operations to a sheet and its native services. */
  public SheetDrawings(Sheet sheet) {
    this(sheet, new DrawingAdapter(), null);
  }

  /** Binds drawing operations to a sheet and its native services. */
  public SheetDrawings(Sheet sheet, DrawingAdapter drawingController) {
    this(sheet, drawingController, null);
  }

  /** Binds drawing operations to a sheet and its native services. */
  public SheetDrawings(
      Sheet sheet, DrawingAdapter drawingController, @Nullable FormulaRuntime formulaRuntime) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.drawingController =
        Objects.requireNonNull(drawingController, "drawingController must not be null");
    this.formulaRuntime = formulaRuntime;
  }

  /** Creates or replaces the named picture. */
  public SheetDrawings setPicture(ResolvedPicture definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setPicture(xssfSheet(), definition);
    return this;
  }

  /** Creates or replaces the named signature line. */
  public SheetDrawings setSignatureLine(ResolvedSignatureLine definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setSignatureLine(xssfSheet(), definition);
    return this;
  }

  /** Creates or replaces the named chart. */
  public SheetDrawings setChart(ResolvedChart definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setChart(xssfSheet(), definition, formulaRuntime);
    return this;
  }

  /** Creates or replaces the named shape. */
  public SheetDrawings setShape(ResolvedShape definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setShape(xssfSheet(), definition);
    return this;
  }

  /** Creates or replaces the named embedded object and preview. */
  public SheetDrawings setEmbeddedObject(ResolvedEmbeddedObject definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setEmbeddedObject(xssfSheet(), definition);
    return this;
  }

  /** Changes the anchor of the selected drawing object. */
  public SheetDrawings setDrawingObjectAnchor(String objectName, DrawingAnchorSpec.TwoCell anchor) {
    TextValidation.requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");
    drawingController.setDrawingObjectAnchor(xssfSheet(), objectName, anchor);
    return this;
  }

  /** Deletes the selected drawing object and unused package relationships. */
  public SheetDrawings deleteDrawingObject(String objectName) {
    TextValidation.requireNonBlank(objectName, "objectName");
    drawingController.deleteDrawingObject(xssfSheet(), objectName);
    return this;
  }

  /** Reads drawing-object facts on the bound sheet. */
  public List<DrawingObjectReport> drawingObjects() {
    return drawingController.drawingObjects(xssfSheet(), formulaRuntime);
  }

  /** Reads chart facts on the bound sheet. */
  public List<ChartReport> charts() {
    return drawingController.charts(xssfSheet(), formulaRuntime);
  }

  /** Reads the selected drawing object payload. */
  public DrawingObjectPayloadReport drawingObjectPayload(String objectName) {
    TextValidation.requireNonBlank(objectName, "objectName");
    return drawingController.drawingObjectPayload(xssfSheet(), objectName);
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }
}
