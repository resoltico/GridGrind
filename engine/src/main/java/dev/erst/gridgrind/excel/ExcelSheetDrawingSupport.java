package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineDefinition;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;

/** Drawing-domain mutations and readback for one sheet wrapper. */
final class ExcelSheetDrawingSupport {
  private final Sheet sheet;
  private final ExcelDrawingController drawingController;
  private final @Nullable ExcelFormulaRuntime formulaRuntime;

  ExcelSheetDrawingSupport(Sheet sheet) {
    this(sheet, new ExcelDrawingController(), null);
  }

  ExcelSheetDrawingSupport(Sheet sheet, ExcelDrawingController drawingController) {
    this(sheet, drawingController, null);
  }

  ExcelSheetDrawingSupport(
      Sheet sheet,
      ExcelDrawingController drawingController,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.drawingController =
        Objects.requireNonNull(drawingController, "drawingController must not be null");
    this.formulaRuntime = formulaRuntime;
  }

  ExcelSheet setPicture(ExcelPictureDefinition definition, ExcelSheet owner) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setPicture(xssfSheet(), definition);
    return owner;
  }

  ExcelSheet setSignatureLine(ExcelSignatureLineDefinition definition, ExcelSheet owner) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setSignatureLine(xssfSheet(), definition);
    return owner;
  }

  ExcelSheet setChart(ExcelChartDefinition definition, ExcelSheet owner) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setChart(xssfSheet(), definition, formulaRuntime);
    return owner;
  }

  ExcelSheet setShape(ExcelShapeDefinition definition, ExcelSheet owner) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setShape(xssfSheet(), definition);
    return owner;
  }

  ExcelSheet setEmbeddedObject(ExcelEmbeddedObjectDefinition definition, ExcelSheet owner) {
    Objects.requireNonNull(definition, "definition must not be null");
    drawingController.setEmbeddedObject(xssfSheet(), definition);
    return owner;
  }

  ExcelSheet setDrawingObjectAnchor(
      String objectName, ExcelDrawingAnchor.TwoCell anchor, ExcelSheet owner) {
    requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");
    drawingController.setDrawingObjectAnchor(xssfSheet(), objectName, anchor);
    return owner;
  }

  ExcelSheet deleteDrawingObject(String objectName, ExcelSheet owner) {
    requireNonBlank(objectName, "objectName");
    drawingController.deleteDrawingObject(xssfSheet(), objectName);
    return owner;
  }

  List<ExcelDrawingObjectSnapshot> drawingObjects() {
    return drawingController.drawingObjects(xssfSheet(), formulaRuntime);
  }

  List<ExcelChartSnapshot> charts() {
    return drawingController.charts(xssfSheet(), formulaRuntime);
  }

  ExcelDrawingObjectPayload drawingObjectPayload(String objectName) {
    requireNonBlank(objectName, "objectName");
    return drawingController.drawingObjectPayload(xssfSheet(), objectName);
  }

  void cleanupEmptyDrawingPatriarch() {
    drawingController.cleanupEmptyDrawingPatriarch(xssfSheet());
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }
}
