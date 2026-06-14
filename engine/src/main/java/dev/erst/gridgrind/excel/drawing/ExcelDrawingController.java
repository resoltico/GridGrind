package dev.erst.gridgrind.excel.drawing;

import dev.erst.gridgrind.excel.DrawingObjectNotFoundException;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelFormulaRuntime;
import dev.erst.gridgrind.excel.ExcelIoSupport;
import dev.erst.gridgrind.excel.ExcelWorkbookImageCatalogSupport;
import dev.erst.gridgrind.excel.PoiRelationRemoval;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlObject;
import org.jspecify.annotations.Nullable;

/** Package-aware drawing and media controller for read and mutation workflows. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelDrawingController {
  private final BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover;
  private final ExcelSignatureLineController signatureLineController;

  public ExcelDrawingController() {
    this(PoiRelationRemoval.defaultRemover());
  }

  public ExcelDrawingController(
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    this(poiRelationRemover, new ExcelSignatureLineController());
  }

  public ExcelDrawingController(
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover,
      ExcelSignatureLineController signatureLineController) {
    this.poiRelationRemover =
        Objects.requireNonNull(poiRelationRemover, "poiRelationRemover must not be null");
    this.signatureLineController =
        Objects.requireNonNull(signatureLineController, "signatureLineController must not be null");
  }

  public List<ExcelDrawingObjectSnapshot> drawingObjects(XSSFSheet sheet) {
    return drawingObjects(sheet, null);
  }

  public List<ExcelDrawingObjectSnapshot> drawingObjects(
      XSSFSheet sheet, @Nullable ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    List<ExcelDrawingObjectSnapshot> snapshots = new ArrayList<>();
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing != null) {
      for (XSSFShape shape : drawing.getShapes()) {
        snapshots.add(snapshot(drawing, shape, formulaRuntime));
      }
    }
    for (ExcelSignatureLineSnapshot signatureLine : signatureLineController.signatureLines(sheet)) {
      snapshots.add(toDrawingObjectSnapshot(signatureLine));
    }
    return List.copyOf(snapshots);
  }

  public List<ExcelChartSnapshot> charts(XSSFSheet sheet) {
    return charts(sheet, null);
  }

  public List<ExcelChartSnapshot> charts(
      XSSFSheet sheet, @Nullable ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return List.of();
    }
    drawing.getShapes();

    List<ExcelChartSnapshot> snapshots = new ArrayList<>();
    for (XSSFChart chart : drawing.getCharts()) {
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();
      if (graphicFrame == null) {
        continue;
      }
      snapshots.add(ExcelDrawingChartSupport.snapshotChart(chart, graphicFrame, formulaRuntime));
    }
    return List.copyOf(snapshots);
  }

  public ExcelDrawingObjectPayload drawingObjectPayload(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelDrawingArgumentSupport.requireNonBlank(objectName, "objectName");

    Optional<LocatedShape> located =
        ExcelDrawingObjectLocator.optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located.isPresent() && signatureLine) {
      throw ExcelDrawingObjectLocator.ambiguousObjectName(sheet, objectName);
    }
    if (located.isEmpty()) {
      if (signatureLine) {
        throw ExcelDrawingObjectLocator.noBinaryPayloadException(sheet, objectName);
      }
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    XSSFShape shape = located.orElseThrow().shape();
    if (shape instanceof XSSFPicture picture) {
      return ExcelDrawingBinarySupport.picturePayload(objectName, picture);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return ExcelDrawingBinarySupport.embeddedObjectPayload(objectName, objectData);
    }
    throw ExcelDrawingObjectLocator.noBinaryPayloadException(sheet, objectName);
  }

  public void setPicture(XSSFSheet sheet, ExcelPictureDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        ExcelDrawingAnchorSupport.toPoiAnchor(drawing, definition.anchor());
    int pictureIndex =
        ExcelWorkbookImageCatalogSupport.addPicture(
            sheet.getWorkbook(),
            definition.imageData().bytes(),
            ExcelPicturePoiBridge.toPoiPictureType(definition.format()));
    XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);
    picture.getCTPicture().getNvPicPr().getCNvPr().setName(definition.name());
    definition
        .description()
        .ifPresent(
            description -> picture.getCTPicture().getNvPicPr().getCNvPr().setDescr(description));
  }

  public void setChart(XSSFSheet sheet, ExcelChartDefinition definition) {
    setChart(sheet, definition, null);
  }

  public void setChart(
      XSSFSheet sheet,
      ExcelChartDefinition definition,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    ExcelDrawingChartSupport.validateChart(sheet, definition, formulaRuntime);
    deleteNamedObjectIfPresent(sheet, definition.name());
    try {
      ExcelDrawingChartSupport.createChart(sheet, definition, formulaRuntime);
    } catch (RuntimeException exception) {
      deleteNamedObjectIfPresent(sheet, definition.name());
      throw exception;
    }
  }

  public void setSignatureLine(XSSFSheet sheet, ExcelSignatureLineDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    signatureLineController.setSignatureLine(sheet, definition);
  }

  public void setShape(XSSFSheet sheet, ExcelShapeDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        ExcelDrawingAnchorSupport.toPoiAnchor(drawing, definition.anchor());
    switch (definition) {
      case ExcelShapeDefinition.SimpleShape simpleShape -> {
        int resolvedShapeType =
            ExcelDrawingAnchorSupport.shapeType(simpleShape.presetGeometryToken());
        org.apache.poi.xssf.usermodel.XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.getCTShape().getNvSpPr().getCNvPr().setName(simpleShape.name());
        shape.setShapeType(resolvedShapeType);
        simpleShape.text().ifPresent(shape::setText);
        return;
      }
      case ExcelShapeDefinition.Connector connector -> {
        org.apache.poi.xssf.usermodel.XSSFConnector poiConnector = drawing.createConnector(anchor);
        poiConnector.getCTConnector().getNvCxnSpPr().getCNvPr().setName(connector.name());
        return;
      }
    }
  }

  public void setEmbeddedObject(XSSFSheet sheet, ExcelEmbeddedObjectDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        ExcelDrawingAnchorSupport.toPoiAnchor(drawing, definition.anchor());
    org.apache.poi.xssf.usermodel.XSSFObjectData objectData =
        ExcelIoSupport.unchecked(
            "Failed to create embedded object '" + definition.name() + "'",
            () -> {
              int storageId =
                  sheet
                      .getWorkbook()
                      .addOlePackage(
                          definition.payload().bytes(),
                          definition.label(),
                          definition.fileName(),
                          definition.command());
              int pictureIndex =
                  ExcelWorkbookImageCatalogSupport.addPicture(
                      sheet.getWorkbook(),
                      definition.previewImage().bytes(),
                      ExcelPicturePoiBridge.toPoiPictureType(definition.previewFormat()));
              return drawing.createObjectData(anchor, storageId, pictureIndex);
            });
    objectData.getCTShape().getNvSpPr().getCNvPr().setName(definition.name());
  }

  public void setDrawingObjectAnchor(
      XSSFSheet sheet, String objectName, ExcelDrawingAnchor.TwoCell anchor) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelDrawingArgumentSupport.requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");

    Optional<LocatedShape> located =
        ExcelDrawingObjectLocator.optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located.isPresent() && signatureLine) {
      throw ExcelDrawingObjectLocator.ambiguousObjectName(sheet, objectName);
    }
    if (signatureLine) {
      signatureLineController.updateAnchorIfPresent(sheet, objectName, anchor);
      return;
    }
    LocatedShape requiredLocated =
        ExcelDrawingObjectLocator.requiredLocatedShape(sheet, objectName);
    if (supportsAnchorMutation(requiredLocated)) {
      ExcelDrawingAnchorSupport.updateAnchorInPlace(
          sheet, objectName, requiredLocated.parentAnchor(), anchor);
      return;
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is read-only until a later parity phase");
  }

  private boolean supportsAnchorMutation(LocatedShape locatedShape) {
    XSSFShape shape = locatedShape.shape();
    if (shape instanceof XSSFPicture
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape) {
      return true;
    }
    return shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame
        && ExcelDrawingChartSupport.chartForGraphicFrame(locatedShape.drawing(), graphicFrame)
            != null;
  }

  public void deleteDrawingObject(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelDrawingArgumentSupport.requireNonBlank(objectName, "objectName");
    Optional<LocatedShape> located =
        ExcelDrawingObjectLocator.optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located.isPresent() && signatureLine) {
      throw ExcelDrawingObjectLocator.ambiguousObjectName(sheet, objectName);
    }
    if (located.isPresent()) {
      deleteLocatedShape(sheet, located.orElseThrow());
      return;
    }
    if (signatureLineController.deleteIfPresent(sheet, objectName)) {
      return;
    }
    throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
  }

  public void cleanupEmptyDrawingPatriarch(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    // POI exposes no supported API for unregistering a live XSSFDrawing relation object from its
    // parent sheet. Hard-deleting the package part leaves the in-memory relation graph stale and
    // breaks later save/commit flows on reopened workbooks. Prefer preserving an inert empty
    // drawing part over emitting a corrupt package.
  }

  private ExcelDrawingObjectSnapshot snapshot(
      XSSFDrawing drawing, XSSFShape shape, @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelDrawingSnapshotSupport.snapshot(drawing, shape, formulaRuntime);
  }

  private ExcelDrawingObjectSnapshot.SignatureLine toDrawingObjectSnapshot(
      ExcelSignatureLineSnapshot signatureLine) {
    return new ExcelDrawingObjectSnapshot.SignatureLine(
        signatureLine.name(),
        signatureLine.anchor(),
        signatureLine.setup(),
        signatureLine.preview());
  }

  private void deleteNamedShapeIfPresent(XSSFSheet sheet, String objectName) {
    Optional<LocatedShape> located =
        ExcelDrawingObjectLocator.optionalLocatedShape(sheet, objectName);
    if (located.isPresent()) {
      deleteLocatedShape(sheet, located.orElseThrow());
    }
  }

  private void deleteNamedObjectIfPresent(XSSFSheet sheet, String objectName) {
    deleteNamedShapeIfPresent(sheet, objectName);
    signatureLineController.deleteIfPresent(sheet, objectName);
  }

  private void deleteLocatedShape(XSSFSheet sheet, LocatedShape located) {
    ExcelDrawingRemovalSupport.deleteLocatedShape(sheet, located, poiRelationRemover);
    cleanupEmptyDrawingPatriarch(sheet);
  }

  public LocatedShape requiredLocatedShape(XSSFSheet sheet, String objectName) {
    return ExcelDrawingObjectLocator.requiredLocatedShape(sheet, objectName);
  }

  /** Supported scalar kinds extracted from chart source cells. */
  public enum CellScalarKind {
    STRING,
    NUMERIC
  }

  public record CellScalar(CellScalarKind kind, @Nullable String text, double number) {
    public CellScalar {
      Objects.requireNonNull(kind, "kind must not be null");
      if (kind == CellScalarKind.STRING) {
        Objects.requireNonNull(text, "text must not be null");
      }
    }
  }

  public record LocatedShape(
      XSSFDrawing drawing,
      XSSFShape shape,
      @Nullable XmlObject shapeXml,
      @Nullable XmlObject parentAnchor) {}
}
