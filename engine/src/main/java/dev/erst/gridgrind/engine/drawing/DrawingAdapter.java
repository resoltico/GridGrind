package dev.erst.gridgrind.engine.drawing;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.engine.backend.xssf.DrawingObjectNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.PoiRelationRemoval;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.packaging.PackageRelations;
import dev.erst.gridgrind.engine.poi.EnumMappings;
import dev.erst.gridgrind.engine.poi.ImageCatalog;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

/** Owns DrawingML objects, pictures, embedded payloads, anchors and relationship-aware deletion. */
// The unexported adapter owns one complete feature; method count does not split its boundary.
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired", "PMD.ExcessivePublicCount"})
public final class DrawingAdapter {

  private final BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover;
  private final SignatureAdapter signatureLineController;

  public DrawingAdapter() {
    this(PoiRelationRemoval.defaultRemover());
  }

  public DrawingAdapter(BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    this(poiRelationRemover, new SignatureAdapter());
  }

  public DrawingAdapter(
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover,
      SignatureAdapter signatureLineController) {
    this.poiRelationRemover =
        Objects.requireNonNull(poiRelationRemover, "poiRelationRemover must not be null");
    this.signatureLineController =
        Objects.requireNonNull(signatureLineController, "signatureLineController must not be null");
  }

  public List<DrawingObjectReport> drawingObjects(XSSFSheet sheet) {
    return drawingObjects(sheet, null);
  }

  public List<DrawingObjectReport> drawingObjects(
      XSSFSheet sheet, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    List<DrawingObjectReport> snapshots = new ArrayList<>();
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing != null) {
      for (XSSFShape shape : drawing.getShapes()) {
        snapshots.add(snapshot(drawing, shape, formulaRuntime));
      }
    }
    for (DrawingObjectReport.SignatureLine signatureLine :
        signatureLineController.signatureLines(sheet)) {
      snapshots.add(signatureLine);
    }
    return List.copyOf(snapshots);
  }

  public List<ChartReport> charts(XSSFSheet sheet) {
    return charts(sheet, null);
  }

  public List<ChartReport> charts(XSSFSheet sheet, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return List.of();
    }
    drawing.getShapes();

    List<ChartReport> snapshots = new ArrayList<>();
    for (XSSFChart chart : drawing.getCharts()) {
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();
      if (graphicFrame == null) {
        continue;
      }
      snapshots.add(ChartAdapter.snapshotChart(chart, graphicFrame, formulaRuntime));
    }
    return List.copyOf(snapshots);
  }

  public DrawingObjectPayloadReport drawingObjectPayload(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");

    Optional<LocatedShape> located = optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located.isPresent() && signatureLine) {
      throw ambiguousObjectName(sheet, objectName);
    }
    if (located.isEmpty()) {
      if (signatureLine) {
        throw noBinaryPayloadException(sheet, objectName);
      }
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    XSSFShape shape = located.orElseThrow().shape();
    if (shape instanceof XSSFPicture picture) {
      return picturePayload(objectName, picture);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return embeddedObjectPayload(objectName, objectData);
    }
    throw noBinaryPayloadException(sheet, objectName);
  }

  public void setPicture(XSSFSheet sheet, ResolvedPicture definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    int pictureIndex =
        ImageCatalog.addPicture(
            sheet.getWorkbook(),
            definition.imageData().bytes(),
            toPoiPictureType(definition.format()));
    XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);
    picture.getCTPicture().getNvPicPr().getCNvPr().setName(definition.name());
    definition
        .description()
        .ifPresent(
            description -> picture.getCTPicture().getNvPicPr().getCNvPr().setDescr(description));
  }

  public void setChart(XSSFSheet sheet, ResolvedChart definition) {
    setChart(sheet, definition, null);
  }

  public void setChart(
      XSSFSheet sheet, ResolvedChart definition, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    ChartAdapter.validateChart(sheet, definition, formulaRuntime);
    deleteNamedObjectIfPresent(sheet, definition.name());
    try {
      ChartAdapter.createChart(sheet, definition, formulaRuntime);
    } catch (RuntimeException exception) {
      deleteNamedObjectIfPresent(sheet, definition.name());
      throw exception;
    }
  }

  public void setSignatureLine(XSSFSheet sheet, ResolvedSignatureLine definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    signatureLineController.setSignatureLine(sheet, definition);
  }

  public void setShape(XSSFSheet sheet, ResolvedShape definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    switch (definition) {
      case ResolvedShape.SimpleShape simpleShape -> {
        int resolvedShapeType = shapeType(simpleShape.presetGeometryToken());
        org.apache.poi.xssf.usermodel.XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.getCTShape().getNvSpPr().getCNvPr().setName(simpleShape.name());
        shape.setShapeType(resolvedShapeType);
        simpleShape.text().ifPresent(shape::setText);
        return;
      }
      case ShapeSpec.Connector connector -> {
        org.apache.poi.xssf.usermodel.XSSFConnector poiConnector = drawing.createConnector(anchor);
        poiConnector.getCTConnector().getNvCxnSpPr().getCNvPr().setName(connector.name());
        return;
      }
    }
  }

  public void setEmbeddedObject(XSSFSheet sheet, ResolvedEmbeddedObject definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    org.apache.poi.xssf.usermodel.XSSFObjectData objectData;
    try {
      int storageId =
          sheet
              .getWorkbook()
              .addOlePackage(
                  definition.payload().bytes(),
                  definition.label(),
                  definition.fileName(),
                  definition.command());
      int pictureIndex =
          ImageCatalog.addPicture(
              sheet.getWorkbook(),
              definition.previewImage().bytes(),
              toPoiPictureType(definition.previewFormat()));
      objectData = drawing.createObjectData(anchor, storageId, pictureIndex);
    } catch (java.io.IOException failure) {
      throw new java.io.UncheckedIOException(
          "Failed to create embedded object '" + definition.name() + "'", failure);
    }
    objectData.getCTShape().getNvSpPr().getCNvPr().setName(definition.name());
  }

  public void setDrawingObjectAnchor(
      XSSFSheet sheet, String objectName, DrawingAnchorSpec.TwoCell anchor) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");

    Optional<LocatedShape> located = optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located.isPresent() && signatureLine) {
      throw ambiguousObjectName(sheet, objectName);
    }
    if (signatureLine) {
      signatureLineController.updateAnchorIfPresent(sheet, objectName, anchor);
      return;
    }
    LocatedShape requiredLocated = requiredLocatedShape(sheet, objectName);
    if (supportsAnchorMutation(requiredLocated)) {
      updateAnchorInPlace(sheet, objectName, requiredLocated.parentAnchor(), anchor);
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
        && ChartAdapter.chartForGraphicFrame(locatedShape.drawing(), graphicFrame) != null;
  }

  public void deleteDrawingObject(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    Optional<LocatedShape> located = optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located.isPresent() && signatureLine) {
      throw ambiguousObjectName(sheet, objectName);
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

  private void deleteNamedShapeIfPresent(XSSFSheet sheet, String objectName) {
    Optional<LocatedShape> located = optionalLocatedShape(sheet, objectName);
    if (located.isPresent()) {
      deleteLocatedShape(sheet, located.orElseThrow());
    }
  }

  private void deleteNamedObjectIfPresent(XSSFSheet sheet, String objectName) {
    deleteNamedShapeIfPresent(sheet, objectName);
    signatureLineController.deleteIfPresent(sheet, objectName);
  }

  private void deleteLocatedShape(XSSFSheet sheet, LocatedShape located) {
    deleteLocatedShape(sheet, located, poiRelationRemover);
  }

  public record LocatedShape(
      XSSFDrawing drawing,
      XSSFShape shape,
      @Nullable XmlObject shapeXml,
      @Nullable XmlObject parentAnchor) {}

  public static org.apache.poi.xssf.usermodel.XSSFClientAnchor toPoiAnchor(
      XSSFDrawing drawing, DrawingAnchorSpec anchor) {
    DrawingAnchorSpec.TwoCell twoCell =
        switch (anchor) {
          case DrawingAnchorSpec.TwoCell value -> value;
        };
    org.apache.poi.xssf.usermodel.XSSFClientAnchor poiAnchor =
        drawing.createAnchor(
            twoCell.from().dx(),
            twoCell.from().dy(),
            twoCell.to().dx(),
            twoCell.to().dy(),
            twoCell.from().columnIndex(),
            twoCell.from().rowIndex(),
            twoCell.to().columnIndex(),
            twoCell.to().rowIndex());
    poiAnchor.setAnchorType(toPoiBehavior(twoCell.behavior()));
    return poiAnchor;
  }

  public static ClientAnchor.AnchorType toPoiBehavior(ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE -> ClientAnchor.AnchorType.MOVE_AND_RESIZE;
      case MOVE_DONT_RESIZE -> ClientAnchor.AnchorType.MOVE_DONT_RESIZE;
      case DONT_MOVE_AND_RESIZE -> ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;
    };
  }

  public static DrawingAnchorReport snapshotAnchor(XmlObject shapeXml) {
    XmlObject parentAnchor = parentAnchor(shapeXml).orElse(null);
    return switch (parentAnchor) {
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor
              twoCellAnchor ->
          new DrawingAnchorReport.TwoCell(
              marker(twoCellAnchor.getFrom()),
              marker(twoCellAnchor.getTo()),
              behavior(twoCellAnchor.getEditAs()));
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor
              oneCellAnchor ->
          new DrawingAnchorReport.OneCell(
              marker(oneCellAnchor.getFrom()),
              oneCellAnchor.getExt().getCx(),
              oneCellAnchor.getExt().getCy(),
              ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor
              absoluteAnchor ->
          new DrawingAnchorReport.Absolute(
              ((Number) absoluteAnchor.getPos().getX()).longValue(),
              ((Number) absoluteAnchor.getPos().getY()).longValue(),
              absoluteAnchor.getExt().getCx(),
              absoluteAnchor.getExt().getCy(),
              ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE);
      case null -> throw new IllegalStateException("Drawing object is missing its parent anchor");
      default ->
          throw new IllegalStateException(
              "Unsupported parent anchor type: " + parentAnchor.getClass().getName());
    };
  }

  public static ExcelDrawingAnchorBehavior behavior(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.Enum editAs) {
    return switch (editAs == null
        ? org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.INT_TWO_CELL
        : editAs.intValue()) {
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.INT_ONE_CELL ->
          ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE;
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.INT_ABSOLUTE ->
          ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE;
      default -> ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE;
    };
  }

  public static DrawingMarkerReport marker(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker marker) {
    return new DrawingMarkerReport(
        marker.getCol(),
        marker.getRow(),
        Math.toIntExact(((Number) marker.getColOff()).longValue()),
        Math.toIntExact(((Number) marker.getRowOff()).longValue()));
  }

  public static XmlObject shapeXml(XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return picture.getCTPicture();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return objectData.getCTShape();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector connector) {
      return connector.getCTConnector();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
      return group.getCTGroupShape();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      return graphicFrame.getCTGraphicalObjectFrame();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape simpleShape) {
      return simpleShape.getCTShape();
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  @SuppressWarnings("PMD.UseTryWithResources")
  public static Optional<XmlObject> parentAnchor(@Nullable XmlObject shapeXml) {
    if (shapeXml == null) {
      return Optional.empty();
    }
    org.apache.xmlbeans.XmlCursor cursor = shapeXml.newCursor();
    try {
      return cursor.toParent() ? Optional.of(cursor.getObject()) : Optional.empty();
    } finally {
      cursor.close();
    }
  }

  public static String resolvedName(XSSFShape shape) {
    return blankAsOptional(shape.getShapeName()).orElseGet(() -> defaultName(shape));
  }

  public static int shapeType(String presetGeometryToken) {
    STShapeType.Enum preset = STShapeType.Enum.forString(presetGeometryToken);
    if (preset == null) {
      throw new IllegalArgumentException("Unsupported presetGeometryToken: " + presetGeometryToken);
    }
    return preset.intValue();
  }

  public static void updateAnchorInPlace(
      XSSFSheet sheet,
      String objectName,
      @Nullable XmlObject parentAnchor,
      DrawingAnchorSpec.TwoCell anchor) {
    if (!(parentAnchor
        instanceof
        org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor
            twoCellAnchor)) {
      throw new IllegalArgumentException(
          "Drawing object '"
              + objectName
              + "' on sheet '"
              + sheet.getSheetName()
              + "' is not backed by a two-cell anchor");
    }
    applyMarker(twoCellAnchor.getFrom(), anchor.from());
    applyMarker(twoCellAnchor.getTo(), anchor.to());
    twoCellAnchor.setEditAs(toPoiEditAs(anchor.behavior()));
  }

  public static String defaultName(XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return "Picture-" + picture.getCTPicture().getNvPicPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return "Object-" + objectData.getCTShape().getNvSpPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector connector) {
      return "Connector-" + connector.getCTConnector().getNvCxnSpPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
      return "Group-" + group.getCTGroupShape().getNvGrpSpPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      return "GraphicFrame-"
          + graphicFrame.getCTGraphicalObjectFrame().getNvGraphicFramePr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape simpleShape) {
      return "Shape-" + simpleShape.getCTShape().getNvSpPr().getCNvPr().getId();
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  private static void applyMarker(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker target,
      DrawingMarkerSpec source) {
    target.setCol(source.columnIndex());
    target.setRow(source.rowIndex());
    target.setColOff(source.dx());
    target.setRowOff(source.dy());
  }

  public static org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.Enum
      toPoiEditAs(ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.TWO_CELL;
      case MOVE_DONT_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.ONE_CELL;
      case DONT_MOVE_AND_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.ABSOLUTE;
    };
  }

  /** Returns the value after enforcing that it is present and not blank. */
  public static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  public static DrawingObjectReport.EmbeddedObject snapshotEmbeddedObject(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    String objectName = resolvedName(objectData);
    EmbeddedObjectReadback readback = embeddedObjectReadback(objectName, objectData);
    return new DrawingObjectReport.EmbeddedObject(
        objectName,
        snapshotAnchor(shapeXml(objectData)),
        readback.packagingKind(),
        readback.label(),
        readback.fileName(),
        readback.command(),
        readback.contentType(),
        readback.payload().size(),
        sha256(readback.payload().bytes()),
        readback.previewFormat(),
        readback.previewImage() == null ? null : (long) readback.previewImage().size(),
        readback.previewImage() == null ? null : sha256(readback.previewImage().bytes()));
  }

  private static DrawingObjectPayloadReport.EmbeddedObject embeddedObjectPayload(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    EmbeddedObjectReadback readback = embeddedObjectReadback(objectName, objectData);
    return new DrawingObjectPayloadReport.EmbeddedObject(
        objectName,
        readback.packagingKind(),
        readback.contentType(),
        readback.fileName(),
        sha256(readback.payload().bytes()),
        java.util.Base64.getEncoder().encodeToString(readback.payload().bytes()),
        readback.label(),
        readback.command());
  }

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> previewImagePart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart =
        previewSheetImagePart(objectData);
    if (previewPart.isPresent()) {
      return previewPart;
    }
    Optional<String> drawingRelationId = previewDrawingRelationId(objectData);
    Optional<org.apache.poi.openxml4j.opc.PackagePart> drawingPreviewPart =
        drawingRelationId.flatMap(
            relationId ->
                relatedInternalPart(objectData.getDrawing().getPackagePart(), relationId));
    return drawingPreviewPart.filter(DrawingAdapter::supportedPreviewImagePart);
  }

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> previewSheetImagePart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    Optional<String> sheetRelationId = previewSheetRelationId(objectData.getOleObject());
    if (sheetRelationId.isEmpty()) {
      return Optional.empty();
    }
    Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart =
        relatedInternalPart(sheetPart(objectData), sheetRelationId.orElseThrow());
    return previewPart.filter(DrawingAdapter::supportedPreviewImagePart);
  }

  private static Optional<String> previewSheetImageRelationId(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return previewImagePart(objectData).isPresent()
        ? previewSheetRelationId(objectData.getOleObject())
        : Optional.empty();
  }

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> relatedInternalPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart, @Nullable String relationshipId) {
    if (relationshipId == null || relationshipId.isBlank()) {
      return Optional.empty();
    }
    org.apache.poi.openxml4j.opc.PackageRelationship relationship =
        sourcePart.getRelationship(relationshipId);
    if (relationship == null
        || relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
      return Optional.empty();
    }
    try {
      return Optional.ofNullable(
          sourcePart
              .getPackage()
              .getPart(
                  org.apache.poi.openxml4j.opc.PackagingURIHelper.createPartName(
                      org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
                          sourcePart.getPartName().getURI(), relationship.getTargetURI()))));
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to resolve related package part", exception);
    }
  }

  public static void cleanupWorkbookImagePartIfUnused(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.@Nullable PackagePartName imagePartName) {
    if (imagePartName == null) {
      return;
    }
    if (!imagePartUsed(workbook, imagePartName)) {
      removeRelationshipsToPart(workbook.getPackagePart(), imagePartName);
      PackageRelations.cleanupPackagePartIfUnused(workbook.getPackage(), imagePartName);
    }
    ImageCatalog.synchronizePictureCatalog(workbook);
  }

  public static void removeOleObject(
      XSSFSheet sheet, org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject target) {
    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObjects oleObjects =
        sheet.getCTWorksheet().getOleObjects();
    if (oleObjects == null) {
      return;
    }
    for (int index = 0; index < oleObjects.sizeOfOleObjectArray(); index++) {
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject candidate =
          oleObjects.getOleObjectArray(index);
      if (candidate.equals(target)) {
        oleObjects.removeOleObject(index);
        if (oleObjects.sizeOfOleObjectArray() == 0) {
          sheet.getCTWorksheet().unsetOleObjects();
        }
        return;
      }
    }
  }

  public static Optional<String> previewSheetRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject) {
    if (!oleObject.isSetObjectPr()) {
      return Optional.empty();
    }
    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTObjectPr objectPr =
        oleObject.getObjectPr();
    return objectPr.isSetId() ? blankAsOptional(objectPr.getId()) : Optional.empty();
  }

  public static void setPreviewSheetRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      String relationId) {
    Objects.requireNonNull(oleObject, "oleObject must not be null");
    Objects.requireNonNull(relationId, "relationId must not be null");
    if (relationId.isBlank()) {
      throw new IllegalArgumentException("relationId must not be blank");
    }
    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTObjectPr objectPr =
        oleObject.isSetObjectPr() ? oleObject.getObjectPr() : oleObject.addNewObjectPr();
    objectPr.setId(relationId);
  }

  public static Optional<String> previewDrawingRelationId(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    if (objectData.getCTShape().getSpPr().getBlipFill() == null
        || objectData.getCTShape().getSpPr().getBlipFill().getBlip() == null) {
      return Optional.empty();
    }
    return blankAsOptional(objectData.getCTShape().getSpPr().getBlipFill().getBlip().getEmbed());
  }

  public static BinaryContent binary(byte[] bytes, String label) {
    Objects.requireNonNull(label, "label must not be null");
    return BinaryContent.readback(bytes);
  }

  public static byte[] partBytes(org.apache.poi.openxml4j.opc.PackagePart part) {
    try {
      return part.getInputStream().readAllBytes();
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to read package part bytes", exception);
    }
  }

  public static String sha256(byte[] bytes) {
    try {
      MessageDigest digest = MessageDigest.getInstance("SHA-256");
      return HexFormat.of().formatHex(digest.digest(bytes));
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 digest is unavailable", exception);
    }
  }

  public static Optional<String> blankAsOptional(@Nullable String value) {
    return value == null || value.isBlank() ? Optional.empty() : Optional.of(value);
  }

  public static Optional<String> firstNonBlank(String first, String second) {
    Optional<String> normalizedFirst = blankAsOptional(first);
    return normalizedFirst.isPresent() ? normalizedFirst : blankAsOptional(second);
  }

  public static boolean looksLikeOle2Storage(byte[] bytes) throws IOException {
    if (bytes.length == 0) {
      return false;
    }
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      return org.apache.poi.poifs.filesystem.FileMagic.valueOf(
              org.apache.poi.poifs.filesystem.FileMagic.prepareToCheckMagic(input))
          == org.apache.poi.poifs.filesystem.FileMagic.OLE2;
    }
  }

  public static String partFileName(org.apache.poi.openxml4j.opc.PackagePart part) {
    String partName = part.getPartName().getName();
    return partName.substring(partName.lastIndexOf('/') + 1);
  }

  public static boolean supportedPreviewImagePart(org.apache.poi.openxml4j.opc.PackagePart part) {
    if (part == null) {
      return false;
    }
    try {
      ExcelPictureFormat.fromContentType(part.getContentType());
      return true;
    } catch (IllegalArgumentException ignored) {
      return false;
    }
  }

  private static EmbeddedObjectReadback embeddedObjectReadback(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart = previewImagePart(objectData);
    BinaryContent previewImage =
        previewPart.isPresent()
            ? binary(partBytes(previewPart.orElseThrow()), "preview image")
            : null;
    ExcelPictureFormat previewFormat =
        previewPart.isPresent()
            ? ExcelPictureFormat.fromContentType(previewPart.orElseThrow().getContentType())
            : null;
    Optional<org.apache.poi.openxml4j.opc.PackagePart> objectPart = oleObjectPart(objectData);
    if (objectPart.isEmpty()) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its package relationship");
    }
    org.apache.poi.openxml4j.opc.PackagePart resolvedObjectPart = objectPart.orElseThrow();
    BinaryContent rawPackage = binary(partBytes(resolvedObjectPart), "embedded object package");
    String contentType = resolvedObjectPart.getContentType();
    String label = null;
    String fileName = partFileName(resolvedObjectPart);
    String command = null;
    ExcelEmbeddedObjectPackagingKind packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    BinaryContent payload = rawPackage;
    try {
      if (looksLikeOle2Storage(rawPackage.bytes())) {
        org.apache.poi.poifs.filesystem.Ole10Native nativeData = ole10Native(rawPackage.bytes());
        payload = BinaryContent.readback(nativeData.getDataBuffer());
        label = firstNonBlank(nativeData.getLabel2(), nativeData.getLabel()).orElse(null);
        fileName = firstNonBlank(nativeData.getFileName2(), nativeData.getFileName()).orElse(null);
        command = firstNonBlank(nativeData.getCommand2(), nativeData.getCommand()).orElse(null);
        packagingKind = ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE;
        contentType = "application/octet-stream";
      }
    } catch (IOException | org.apache.poi.poifs.filesystem.Ole10NativeException exception) {
      payload = rawPackage;
      packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    }
    fileName = Objects.requireNonNullElse(fileName, objectName + ".bin");
    return new EmbeddedObjectReadback(
        packagingKind, label, fileName, command, contentType, payload, previewFormat, previewImage);
  }

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> oleObjectPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return relatedInternalPart(sheetPart(objectData), objectData.getOleObject().getId());
  }

  public static boolean imagePartUsed(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    SignatureAdapter signatureLineController = new SignatureAdapter();
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (sheetUsesImagePart(
          workbook.getSheetAt(sheetIndex), imagePartName, signatureLineController)) {
        return true;
      }
    }
    return false;
  }

  private static boolean sheetUsesImagePart(
      XSSFSheet sheet,
      org.apache.poi.openxml4j.opc.PackagePartName imagePartName,
      SignatureAdapter signatureLineController) {
    org.apache.poi.xssf.usermodel.XSSFDrawing drawing = sheet.getDrawingPatriarch();
    return (drawing != null && drawingUsesImagePart(drawing, imagePartName))
        || signatureLineController.usesImagePart(sheet, imagePartName);
  }

  private static boolean drawingUsesImagePart(
      org.apache.poi.xssf.usermodel.XSSFDrawing drawing,
      org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    for (org.apache.poi.xssf.usermodel.XSSFShape shape : drawing.getShapes()) {
      if (shapeUsesImagePart(shape, imagePartName)) {
        return true;
      }
    }
    return false;
  }

  private static boolean shapeUsesImagePart(
      org.apache.poi.xssf.usermodel.XSSFShape shape,
      org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    if (shape instanceof XSSFPicture picture) {
      return imagePartName.equals(imagePartNameOrNull(picture).orElse(null));
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart = previewImagePart(objectData);
      return previewPart.isPresent()
          && previewPart.orElseThrow().getPartName().equals(imagePartName);
    }
    return false;
  }

  public static void removeRelationshipsToPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart,
      org.apache.poi.openxml4j.opc.PackagePartName targetPartName) {
    if (sourcePart.isRelationshipPart()) {
      return;
    }
    List<String> relationshipIds = new ArrayList<>();
    try {
      for (org.apache.poi.openxml4j.opc.PackageRelationship relationship :
          sourcePart.getRelationships()) {
        if (relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
          continue;
        }
        if (targetPartName
            .getURI()
            .equals(
                org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
                    sourcePart.getPartName().getURI(), relationship.getTargetURI()))) {
          relationshipIds.add(relationship.getId());
        }
      }
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to inspect package relationships", exception);
    }
    for (String relationshipId : relationshipIds) {
      sourcePart.removeRelationship(relationshipId);
    }
  }

  private static org.apache.poi.openxml4j.opc.PackagePart sheetPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ((XSSFSheet) objectData.getDrawing().getParent()).getPackagePart();
  }

  private static org.apache.poi.poifs.filesystem.Ole10Native ole10Native(byte[] bytes)
      throws IOException, org.apache.poi.poifs.filesystem.Ole10NativeException {
    try (org.apache.poi.poifs.filesystem.POIFSFileSystem filesystem =
        new org.apache.poi.poifs.filesystem.POIFSFileSystem(new ByteArrayInputStream(bytes))) {
      org.apache.poi.poifs.filesystem.DirectoryNode directory = filesystem.getRoot();
      return org.apache.poi.poifs.filesystem.Ole10Native.createFromEmbeddedOleObject(directory);
    }
  }

  private record EmbeddedObjectReadback(
      ExcelEmbeddedObjectPackagingKind packagingKind,
      @Nullable String label,
      String fileName,
      @Nullable String command,
      String contentType,
      BinaryContent payload,
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable BinaryContent previewImage) {}

  public static LocatedShape requiredLocatedShape(XSSFSheet sheet, String objectName) {
    Optional<LocatedShape> located = optionalLocatedShape(sheet, objectName);
    if (located.isEmpty()) {
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    return located.orElseThrow();
  }

  private static Optional<LocatedShape> optionalLocatedShape(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return Optional.empty();
    }
    XSSFShape matchedShape = null;
    XmlObject matchedShapeXml = null;
    for (XSSFShape shape : drawing.getShapes()) {
      if (!resolvedName(shape).equals(objectName)) {
        continue;
      }
      if (matchedShape != null) {
        throw ambiguousObjectName(sheet, objectName);
      }
      matchedShape = shape;
      matchedShapeXml = shapeXml(shape);
    }
    return matchedShape == null
        ? Optional.empty()
        : Optional.of(
            new LocatedShape(
                drawing,
                matchedShape,
                matchedShapeXml,
                parentAnchor(matchedShapeXml).orElse(null)));
  }

  private static IllegalArgumentException ambiguousObjectName(XSSFSheet sheet, String objectName) {
    return new IllegalArgumentException(
        "Multiple drawing objects named '"
            + objectName
            + "' exist on sheet '"
            + sheet.getSheetName()
            + "'");
  }

  private static IllegalArgumentException noBinaryPayloadException(
      XSSFSheet sheet, String objectName) {
    return new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' does not expose a binary payload");
  }

  public static DrawingObjectReport.Picture snapshotPicture(XSSFPicture picture) {
    PictureReadback readback = requiredPictureReadback(picture);
    RasterDimensions dimensions = rasterDimensions(readback.bytes());
    return new DrawingObjectReport.Picture(
        resolvedName(picture),
        snapshotAnchor(shapeXml(picture)),
        readback.format(),
        readback.contentType(),
        readback.bytes().length,
        sha256(readback.bytes()),
        dimensions.widthPixels(),
        dimensions.heightPixels(),
        blankAsOptional(picture.getCTPicture().getNvPicPr().getCNvPr().getDescr()).orElse(null));
  }

  public static DrawingObjectPayloadReport.Picture picturePayload(
      String objectName, XSSFPicture picture) {
    PictureReadback readback = requiredPictureReadback(picture);
    return new DrawingObjectPayloadReport.Picture(
        objectName,
        readback.format(),
        readback.contentType(),
        objectName + readback.format().defaultExtension(),
        sha256(readback.bytes()),
        java.util.Base64.getEncoder().encodeToString(readback.bytes()),
        blankAsOptional(picture.getCTPicture().getNvPicPr().getCNvPr().getDescr()).orElse(null));
  }

  public static void deletePicture(XSSFSheet sheet, LocatedShape located, XSSFPicture picture) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(located, "located must not be null");
    Objects.requireNonNull(picture, "picture must not be null");

    PackagePartName imagePartName = imagePartNameOrNull(picture).orElse(null);
    Optional<String> relationId = pictureRelationId(picture);
    if (relationId.isPresent()) {
      located.drawing().getPackagePart().removeRelationship(relationId.orElseThrow());
    }
    removeParentAnchor(located.drawing(), located.parentAnchor());
    if (imagePartName != null) {
      cleanupWorkbookImagePartIfUnused(sheet.getWorkbook(), imagePartName);
    }
  }

  public static PictureReadback requiredPictureReadback(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    XSSFPictureData pictureData = pictureDataOrNull(picture).orElse(null);
    if (pictureData == null) {
      throw missingPictureRelationship(picture);
    }
    byte[] bytes = pictureData.getData();
    return new PictureReadback(
        pictureData,
        pictureData.getPackagePart(),
        fromPoiPictureType(pictureData.getPictureType()),
        pictureData.getPackagePart().getContentType(),
        bytes);
  }

  public static Optional<XSSFPictureData> pictureDataOrNull(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    String relationId = pictureRelationId(picture).orElse(null);
    if (relationId == null) {
      return Optional.empty();
    }
    XSSFPictureData pictureData = picture.getPictureData();
    if (pictureData != null) {
      return Optional.of(pictureData);
    }
    PackagePart imagePart =
        relatedInternalPart(picture.getDrawing().getPackagePart(), relationId).orElse(null);
    if (imagePart == null) {
      return Optional.empty();
    }
    try {
      ExcelPictureFormat.fromContentType(imagePart.getContentType());
    } catch (IllegalArgumentException ignored) {
      return Optional.empty();
    }
    return Optional.of(ImageCatalog.pictureDataForPart(imagePart));
  }

  public static Optional<PackagePart> relatedImagePartOrNull(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    String relationId = pictureRelationId(picture).orElse(null);
    if (relationId == null) {
      return Optional.empty();
    }
    XSSFPictureData pictureData = picture.getPictureData();
    if (pictureData != null) {
      return Optional.of(pictureData.getPackagePart());
    }
    return relatedInternalPart(picture.getDrawing().getPackagePart(), relationId);
  }

  public static Optional<PackagePartName> imagePartNameOrNull(XSSFPicture picture) {
    PackagePart imagePart = relatedImagePartOrNull(picture).orElse(null);
    if (imagePart == null) {
      return Optional.empty();
    }
    try {
      ExcelPictureFormat.fromContentType(imagePart.getContentType());
      return Optional.of(imagePart.getPartName());
    } catch (IllegalArgumentException ignored) {
      return Optional.empty();
    }
  }

  public static Optional<String> pictureRelationId(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    if (picture.getCTPicture().getBlipFill() == null
        || picture.getCTPicture().getBlipFill().getBlip() == null) {
      return Optional.empty();
    }
    return blankAsOptional(picture.getCTPicture().getBlipFill().getBlip().getEmbed());
  }

  public static void setPictureRelationId(XSSFPicture picture, String relationId) {
    Objects.requireNonNull(picture, "picture must not be null");
    Objects.requireNonNull(relationId, "relationId must not be null");
    if (relationId.isBlank()) {
      throw new IllegalArgumentException("relationId must not be blank");
    }
    if (picture.getCTPicture().getBlipFill() == null
        || picture.getCTPicture().getBlipFill().getBlip() == null) {
      throw new IllegalStateException(
          "Picture '"
              + resolvedName(picture)
              + "' on sheet '"
              + sheetName(picture)
              + "' is missing its blip payload");
    }
    picture.getCTPicture().getBlipFill().getBlip().setEmbed(relationId);
  }

  public static Optional<String> reusableRelationId(
      org.apache.poi.xssf.usermodel.XSSFDrawing drawing,
      @Nullable String relationId,
      PackagePartName targetPartName) {
    Objects.requireNonNull(drawing, "drawing must not be null");
    Objects.requireNonNull(targetPartName, "targetPartName must not be null");
    if (relationId == null || relationId.isBlank()) {
      return Optional.empty();
    }
    RelationPart existingRelation = drawing.getRelationPartById(relationId);
    if (existingRelation == null) {
      return Optional.of(relationId);
    }
    return existingRelation.getDocumentPart().getPackagePart().getPartName().equals(targetPartName)
        ? Optional.of(relationId)
        : Optional.empty();
  }

  public static IllegalStateException missingPictureRelationship(XSSFPicture picture) {
    String relationId = pictureRelationId(picture).orElse(null);
    String relationFragment =
        relationId == null ? "its image relationship" : "image relationship '" + relationId + "'";
    return new IllegalStateException(
        "Picture '"
            + resolvedName(picture)
            + "' on sheet '"
            + sheetName(picture)
            + "' is missing "
            + relationFragment);
  }

  private static String sheetName(XSSFPicture picture) {
    return ((XSSFSheet) picture.getDrawing().getParent()).getSheetName();
  }

  /** Resolved picture binary plus enough metadata to rebuild factual readback snapshots. */
  public static final class PictureReadback {
    private final XSSFPictureData pictureData;
    private final PackagePart picturePart;
    private final ExcelPictureFormat format;
    private final String contentType;
    private final byte[] bytes;

    public PictureReadback(
        XSSFPictureData pictureData,
        PackagePart picturePart,
        ExcelPictureFormat format,
        String contentType,
        byte[] bytes) {
      Objects.requireNonNull(pictureData, "pictureData must not be null");
      Objects.requireNonNull(picturePart, "picturePart must not be null");
      Objects.requireNonNull(format, "format must not be null");
      Objects.requireNonNull(contentType, "contentType must not be null");
      if (contentType.isBlank()) {
        throw new IllegalArgumentException("contentType must not be blank");
      }
      this.pictureData = pictureData;
      this.picturePart = picturePart;
      this.format = format;
      this.contentType = contentType;
      this.bytes = Objects.requireNonNull(bytes, "bytes must not be null").clone();
    }

    public XSSFPictureData pictureData() {
      return pictureData;
    }

    public PackagePart picturePart() {
      return picturePart;
    }

    public ExcelPictureFormat format() {
      return format;
    }

    public String contentType() {
      return contentType;
    }

    public byte[] bytes() {
      return bytes.clone();
    }
  }

  public static void deleteLocatedShape(
      XSSFSheet sheet,
      LocatedShape located,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture picture) {
      deletePicture(sheet, located, picture);
      return;
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      XSSFChart chart = ChartAdapter.chartForGraphicFrame(located.drawing(), graphicFrame);
      if (chart != null) {
        deleteChart(located, chart, poiRelationRemover);
        return;
      }
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      deleteEmbeddedObject(sheet, located, objectData);
      return;
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup) {
      removeParentAnchor(located.drawing(), located.parentAnchor());
      return;
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + resolvedName(located.shape())
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is read-only until a later parity phase");
  }

  private static void deleteChart(
      LocatedShape located,
      XSSFChart chart,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    if (!poiRelationRemover.test(located.drawing(), chart)) {
      throw new IllegalStateException(
          "Failed to remove chart relation for '"
              + resolvedName((org.apache.poi.xssf.usermodel.XSSFGraphicFrame) located.shape())
              + "'");
    }
    removeParentAnchor(located.drawing(), located.parentAnchor());
  }

  private static void deleteEmbeddedObject(
      XSSFSheet sheet,
      LocatedShape located,
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    org.apache.poi.openxml4j.opc.PackagePart objectPart =
        relatedInternalPart(sheetPart(objectData), objectData.getOleObject().getId()).orElse(null);
    org.apache.poi.openxml4j.opc.PackagePart previewPart =
        previewImagePart(objectData).orElse(null);
    org.apache.poi.openxml4j.opc.PackagePartName olePartName =
        objectPart == null ? null : objectPart.getPartName();
    org.apache.poi.openxml4j.opc.PackagePartName previewPartName =
        previewPart == null ? null : previewPart.getPartName();
    String drawingPreviewRelationId = previewDrawingRelationId(objectData).orElse(null);
    String sheetPreviewRelationId = previewSheetImageRelationId(objectData).orElse(null);
    String oleRelationId = objectData.getOleObject().getId();

    if (drawingPreviewRelationId != null) {
      located.drawing().getPackagePart().removeRelationship(drawingPreviewRelationId);
    }
    if (sheetPreviewRelationId != null) {
      sheet.getPackagePart().removeRelationship(sheetPreviewRelationId);
    }
    if (oleRelationId != null) {
      sheet.getPackagePart().removeRelationship(oleRelationId);
    }
    removeOleObject(sheet, objectData.getOleObject());
    removeParentAnchor(located.drawing(), located.parentAnchor());
    PackageRelations.cleanupPackagePartIfUnused(sheet.getWorkbook().getPackage(), olePartName);
    cleanupWorkbookImagePartIfUnused(sheet.getWorkbook(), previewPartName);
  }

  public static void removeParentAnchor(XSSFDrawing drawing, @Nullable XmlObject parentAnchor) {
    switch (parentAnchor) {
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor
              twoCellAnchor ->
          removeTwoCellAnchor(drawing.getCTDrawing(), twoCellAnchor);
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor
              oneCellAnchor ->
          removeOneCellAnchor(drawing.getCTDrawing(), oneCellAnchor);
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor
              absoluteAnchor ->
          removeAbsoluteAnchor(drawing.getCTDrawing(), absoluteAnchor);
      case null -> throw new IllegalStateException("Drawing object is missing its parent anchor");
      default ->
          throw new IllegalStateException(
              "Unsupported parent anchor type: " + parentAnchor.getClass().getName());
    }
  }

  public static void removeTwoCellAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor target) {
    for (int index = 0; index < drawing.sizeOfTwoCellAnchorArray(); index++) {
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor candidate =
          drawing.getTwoCellAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeTwoCellAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate two-cell drawing anchor for deletion");
  }

  public static void removeOneCellAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor target) {
    for (int index = 0; index < drawing.sizeOfOneCellAnchorArray(); index++) {
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor candidate =
          drawing.getOneCellAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeOneCellAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate one-cell drawing anchor for deletion");
  }

  public static void removeAbsoluteAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor target) {
    for (int index = 0; index < drawing.sizeOfAbsoluteAnchorArray(); index++) {
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor candidate =
          drawing.getAbsoluteAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeAbsoluteAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate absolute drawing anchor for deletion");
  }

  public static DrawingObjectReport snapshot(XSSFDrawing drawing, XSSFShape shape) {
    return snapshot(drawing, shape, null);
  }

  public static DrawingObjectReport snapshot(
      XSSFDrawing drawing, XSSFShape shape, @Nullable FormulaRuntime formulaRuntime) {
    if (shape instanceof XSSFPicture picture) {
      return snapshotPicture(picture);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return snapshotEmbeddedObject(objectData);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector connector) {
      return snapshotConnector(connector);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
      return snapshotGroup(drawing, group);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      XSSFChart chart = ChartAdapter.chartForGraphicFrame(drawing, graphicFrame);
      return chart == null
          ? snapshotGraphicFrame(graphicFrame)
          : ChartAdapter.snapshotChartDrawingObject(chart, graphicFrame, formulaRuntime);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape simpleShape) {
      return snapshotShape(simpleShape);
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  public static DrawingObjectReport.Shape snapshotShape(
      org.apache.poi.xssf.usermodel.XSSFSimpleShape shape) {
    STShapeType.Enum preset =
        !shape.getCTShape().getSpPr().isSetPrstGeom()
            ? null
            : shape.getCTShape().getSpPr().getPrstGeom().getPrst();
    return new DrawingObjectReport.Shape(
        resolvedName(shape),
        snapshotAnchor(shapeXml(shape)),
        ExcelDrawingShapeKind.SIMPLE_SHAPE,
        preset == null ? null : preset.toString(),
        blankAsOptional(shape.getText()).orElse(null),
        0);
  }

  private static DrawingObjectReport.Shape snapshotConnector(
      org.apache.poi.xssf.usermodel.XSSFConnector connector) {
    return new DrawingObjectReport.Shape(
        resolvedName(connector),
        snapshotAnchor(shapeXml(connector)),
        ExcelDrawingShapeKind.CONNECTOR,
        null,
        null,
        0);
  }

  private static DrawingObjectReport.Shape snapshotGroup(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
    return new DrawingObjectReport.Shape(
        resolvedName(group),
        snapshotAnchor(shapeXml(group)),
        ExcelDrawingShapeKind.GROUP,
        null,
        null,
        drawing.getShapes(group).size());
  }

  public static DrawingObjectReport.Shape snapshotGraphicFrame(
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return new DrawingObjectReport.Shape(
        resolvedName(graphicFrame),
        snapshotAnchor(shapeXml(graphicFrame)),
        ExcelDrawingShapeKind.GRAPHIC_FRAME,
        null,
        null,
        0);
  }

  public static RasterDimensions rasterDimensions(byte[] bytes) {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      BufferedImage image = javax.imageio.ImageIO.read(input);
      return image == null
          ? RasterDimensions.none()
          : new RasterDimensions(image.getWidth(), image.getHeight());
    } catch (IOException exception) {
      return RasterDimensions.none();
    }
  }

  public record RasterDimensions(@Nullable Integer widthPixels, @Nullable Integer heightPixels) {
    public static RasterDimensions none() {
      return new RasterDimensions(null, null);
    }
  }

  private static final Map<ExcelPictureFormat, Integer> TO_POI_PICTURE_TYPE =
      EnumMappings.exactEnumMap(
          ExcelPictureFormat.class,
          "Apache POI picture-type mapping",
          Map.ofEntries(
              Map.entry(ExcelPictureFormat.EMF, Workbook.PICTURE_TYPE_EMF),
              Map.entry(ExcelPictureFormat.WMF, Workbook.PICTURE_TYPE_WMF),
              Map.entry(ExcelPictureFormat.PICT, Workbook.PICTURE_TYPE_PICT),
              Map.entry(ExcelPictureFormat.JPEG, Workbook.PICTURE_TYPE_JPEG),
              Map.entry(ExcelPictureFormat.PNG, Workbook.PICTURE_TYPE_PNG),
              Map.entry(ExcelPictureFormat.DIB, Workbook.PICTURE_TYPE_DIB),
              Map.entry(ExcelPictureFormat.GIF, XSSFWorkbook.PICTURE_TYPE_GIF),
              Map.entry(ExcelPictureFormat.TIFF, XSSFWorkbook.PICTURE_TYPE_TIFF),
              Map.entry(ExcelPictureFormat.EPS, XSSFWorkbook.PICTURE_TYPE_EPS),
              Map.entry(ExcelPictureFormat.BMP, XSSFWorkbook.PICTURE_TYPE_BMP),
              Map.entry(ExcelPictureFormat.WPG, XSSFWorkbook.PICTURE_TYPE_WPG)));

  private static final Map<Integer, ExcelPictureFormat> FROM_POI_PICTURE_TYPE =
      Map.ofEntries(
          Map.entry(Workbook.PICTURE_TYPE_EMF, ExcelPictureFormat.EMF),
          Map.entry(Workbook.PICTURE_TYPE_WMF, ExcelPictureFormat.WMF),
          Map.entry(Workbook.PICTURE_TYPE_PICT, ExcelPictureFormat.PICT),
          Map.entry(Workbook.PICTURE_TYPE_JPEG, ExcelPictureFormat.JPEG),
          Map.entry(Workbook.PICTURE_TYPE_PNG, ExcelPictureFormat.PNG),
          Map.entry(Workbook.PICTURE_TYPE_DIB, ExcelPictureFormat.DIB),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_GIF, ExcelPictureFormat.GIF),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_TIFF, ExcelPictureFormat.TIFF),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_EPS, ExcelPictureFormat.EPS),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_BMP, ExcelPictureFormat.BMP),
          Map.entry(XSSFWorkbook.PICTURE_TYPE_WPG, ExcelPictureFormat.WPG));

  public static int toPoiPictureType(ExcelPictureFormat format) {
    return EnumMappings.requireMappedValue(TO_POI_PICTURE_TYPE, format, "GridGrind picture format");
  }

  public static ExcelPictureFormat fromPoiPictureType(int poiPictureType) {
    ExcelPictureFormat resolved = FROM_POI_PICTURE_TYPE.get(poiPictureType);
    if (resolved == null) {
      throw new IllegalArgumentException("Unsupported POI picture type: " + poiPictureType);
    }
    return resolved;
  }
}
