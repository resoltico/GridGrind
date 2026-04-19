package dev.erst.gridgrind.excel;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Objects;
import java.util.function.BiPredicate;
import java.util.function.BooleanSupplier;
import javax.imageio.ImageIO;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

/** Package-aware drawing and media controller for read and mutation workflows. */
final class ExcelDrawingController {
  private static final javax.xml.namespace.QName OLE_PREVIEW_ID =
      new javax.xml.namespace.QName(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376_NS, "id");
  private final BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover;

  ExcelDrawingController() {
    this(PoiRelationRemoval.defaultRemover());
  }

  ExcelDrawingController(BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    this.poiRelationRemover =
        Objects.requireNonNull(poiRelationRemover, "poiRelationRemover must not be null");
  }

  List<ExcelDrawingObjectSnapshot> drawingObjects(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return List.of();
    }

    List<ExcelDrawingObjectSnapshot> snapshots = new ArrayList<>();
    for (XSSFShape shape : drawing.getShapes()) {
      snapshots.add(snapshot(drawing, shape));
    }
    return List.copyOf(snapshots);
  }

  List<ExcelChartSnapshot> charts(XSSFSheet sheet) {
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
      snapshots.add(snapshotChart(chart, graphicFrame));
    }
    return List.copyOf(snapshots);
  }

  ExcelDrawingObjectPayload drawingObjectPayload(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");

    LocatedShape located = requiredLocatedShape(sheet, objectName);
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture picture) {
      return picturePayload(objectName, picture);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return embeddedObjectPayload(objectName, objectData);
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' does not expose a binary payload");
  }

  void setPicture(XSSFSheet sheet, ExcelPictureDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedShapeIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    int pictureIndex =
        sheet
            .getWorkbook()
            .addPicture(definition.imageData().bytes(), definition.format().poiPictureType());
    XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);
    picture.getCTPicture().getNvPicPr().getCNvPr().setName(definition.name());
    if (definition.description() != null) {
      picture.getCTPicture().getNvPicPr().getCNvPr().setDescr(definition.description());
    }
  }

  void setChart(XSSFSheet sheet, ExcelChartDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    PreparedChartDefinition prepared = prepareChartDefinition(sheet, definition);
    LocatedShape located = optionalLocatedShape(sheet, prepared.name());
    if (located == null) {
      createChart(sheet, prepared);
      return;
    }

    if (located.shape() instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      XSSFChart chart = chartForGraphicFrame(located.drawing(), graphicFrame);
      if (chart != null) {
        mutateChart(sheet, located, chart, prepared);
        return;
      }
    }

    deleteLocatedShape(sheet, located);
    createChart(sheet, prepared);
  }

  void setShape(XSSFSheet sheet, ExcelShapeDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    Integer preparedShapeType =
        definition.kind() == ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE
            ? shapeType(definition.presetGeometryToken())
            : null;
    deleteNamedShapeIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    if (definition.kind() == ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE) {
      org.apache.poi.xssf.usermodel.XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
      shape.getCTShape().getNvSpPr().getCNvPr().setName(definition.name());
      shape.setShapeType(Objects.requireNonNull(preparedShapeType));
      if (definition.text() != null) {
        shape.setText(definition.text());
      }
      return;
    }
    org.apache.poi.xssf.usermodel.XSSFConnector connector = drawing.createConnector(anchor);
    connector.getCTConnector().getNvCxnSpPr().getCNvPr().setName(definition.name());
  }

  void setEmbeddedObject(XSSFSheet sheet, ExcelEmbeddedObjectDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedShapeIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
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
                  sheet
                      .getWorkbook()
                      .addPicture(
                          definition.previewImage().bytes(),
                          definition.previewFormat().poiPictureType());
              return drawing.createObjectData(anchor, storageId, pictureIndex);
            });
    objectData.getCTShape().getNvSpPr().getCNvPr().setName(definition.name());
  }

  void setDrawingObjectAnchor(
      XSSFSheet sheet, String objectName, ExcelDrawingAnchor.TwoCell anchor) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");

    LocatedShape located = requiredLocatedShape(sheet, objectName);
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture
        || (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame
            && chartForGraphicFrame(located.drawing(), graphicFrame) != null)
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape) {
      updateAnchorInPlace(sheet, objectName, located.parentAnchor(), anchor);
      return;
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is read-only until a later parity phase");
  }

  void deleteDrawingObject(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    deleteLocatedShape(sheet, requiredLocatedShape(sheet, objectName));
  }

  void cleanupEmptyDrawingPatriarch(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    // POI exposes no supported API for unregistering a live XSSFDrawing relation object from its
    // parent sheet. Hard-deleting the package part leaves the in-memory relation graph stale and
    // breaks later save/commit flows on reopened workbooks. Prefer preserving an inert empty
    // drawing part over emitting a corrupt package.
  }

  private ExcelDrawingObjectSnapshot snapshot(XSSFDrawing drawing, XSSFShape shape) {
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
      XSSFChart chart = chartForGraphicFrame(drawing, graphicFrame);
      return chart == null
          ? snapshotGraphicFrame(graphicFrame)
          : snapshotChartDrawingObject(chart, graphicFrame);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape simpleShape) {
      return snapshotShape(simpleShape);
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  private ExcelDrawingObjectSnapshot.Picture snapshotPicture(XSSFPicture picture) {
    org.apache.poi.xssf.usermodel.XSSFPictureData pictureData = picture.getPictureData();
    byte[] bytes = pictureData.getData();
    String description = nullIfBlank(picture.getCTPicture().getNvPicPr().getCNvPr().getDescr());
    RasterDimensions dimensions = rasterDimensions(bytes);
    return new ExcelDrawingObjectSnapshot.Picture(
        resolvedName(picture),
        snapshotAnchor(shapeXml(picture)),
        ExcelPictureFormat.fromPoiPictureType(pictureData.getPictureType()),
        pictureData.getPackagePart().getContentType(),
        bytes.length,
        sha256(bytes),
        dimensions.widthPixels(),
        dimensions.heightPixels(),
        description);
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotConnector(
      org.apache.poi.xssf.usermodel.XSSFConnector connector) {
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(connector),
        snapshotAnchor(shapeXml(connector)),
        ExcelDrawingShapeKind.CONNECTOR,
        null,
        null,
        0);
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotGroup(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(group),
        snapshotAnchor(shapeXml(group)),
        ExcelDrawingShapeKind.GROUP,
        null,
        null,
        drawing.getShapes(group).size());
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotGraphicFrame(
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(graphicFrame),
        snapshotAnchor(shapeXml(graphicFrame)),
        ExcelDrawingShapeKind.GRAPHIC_FRAME,
        null,
        null,
        0);
  }

  private ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    ExcelChartSnapshot snapshot = snapshotChart(chart, graphicFrame);
    return new ExcelDrawingObjectSnapshot.Chart(
        resolvedChartName(graphicFrame),
        snapshotAnchor(shapeXml(graphicFrame)),
        !(snapshot instanceof ExcelChartSnapshot.Unsupported),
        chartPlotTypeTokens(chart),
        titleSummary(snapshot));
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotShape(
      org.apache.poi.xssf.usermodel.XSSFSimpleShape shape) {
    STShapeType.Enum preset =
        !shape.getCTShape().getSpPr().isSetPrstGeom()
            ? null
            : shape.getCTShape().getSpPr().getPrstGeom().getPrst();
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(shape),
        snapshotAnchor(shapeXml(shape)),
        ExcelDrawingShapeKind.SIMPLE_SHAPE,
        preset == null ? null : preset.toString(),
        nullIfBlank(shape.getText()),
        0);
  }

  private ExcelChartSnapshot snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    List<org.apache.poi.xddf.usermodel.chart.XDDFChartData> chartData = chart.getChartSeries();
    List<String> plotTypeTokens = chartPlotTypeTokens(chartData);
    if (chartData.size() != 1) {
      return new ExcelChartSnapshot.Unsupported(
          resolvedChartName(graphicFrame),
          snapshotAnchor(shapeXml(graphicFrame)),
          plotTypeTokens,
          "Only single-plot simple charts are modeled authoritatively.");
    }

    org.apache.poi.xddf.usermodel.chart.XDDFChartData data = chartData.getFirst();
    try {
      return switch (data) {
        case org.apache.poi.xddf.usermodel.chart.XDDFBarChartData barChartData ->
            new ExcelChartSnapshot.Bar(
                resolvedChartName(graphicFrame),
                snapshotAnchor(shapeXml(graphicFrame)),
                snapshotTitle(chart),
                snapshotLegend(chart),
                snapshotDisplayBlanks(chart),
                chart.isPlotOnlyVisibleCells(),
                barVaryColors(chart),
                ExcelChartPoiBridge.fromPoiBarDirection(barChartData.getBarDirection()),
                snapshotAxes(chart),
                snapshotSeries(barChartData));
        case org.apache.poi.xddf.usermodel.chart.XDDFLineChartData lineChartData ->
            new ExcelChartSnapshot.Line(
                resolvedChartName(graphicFrame),
                snapshotAnchor(shapeXml(graphicFrame)),
                snapshotTitle(chart),
                snapshotLegend(chart),
                snapshotDisplayBlanks(chart),
                chart.isPlotOnlyVisibleCells(),
                lineVaryColors(chart),
                snapshotAxes(chart),
                snapshotSeries(lineChartData));
        case org.apache.poi.xddf.usermodel.chart.XDDFPieChartData pieChartData ->
            new ExcelChartSnapshot.Pie(
                resolvedChartName(graphicFrame),
                snapshotAnchor(shapeXml(graphicFrame)),
                snapshotTitle(chart),
                snapshotLegend(chart),
                snapshotDisplayBlanks(chart),
                chart.isPlotOnlyVisibleCells(),
                pieVaryColors(chart),
                pieChartData.getFirstSliceAngle(),
                snapshotSeries(pieChartData));
        default ->
            new ExcelChartSnapshot.Unsupported(
                resolvedChartName(graphicFrame),
                snapshotAnchor(shapeXml(graphicFrame)),
                plotTypeTokens,
                "Chart plot family is outside the current modeled simple-chart contract.");
      };
    } catch (IllegalArgumentException | IllegalStateException exception) {
      return new ExcelChartSnapshot.Unsupported(
          resolvedChartName(graphicFrame),
          snapshotAnchor(shapeXml(graphicFrame)),
          plotTypeTokens,
          exception.getMessage());
    }
  }

  private ExcelDrawingObjectSnapshot.EmbeddedObject snapshotEmbeddedObject(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    EmbeddedObjectReadback readback = embeddedObjectReadback(resolvedName(objectData), objectData);
    return new ExcelDrawingObjectSnapshot.EmbeddedObject(
        resolvedName(objectData),
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

  private ExcelDrawingObjectPayload.Picture picturePayload(String objectName, XSSFPicture picture) {
    org.apache.poi.xssf.usermodel.XSSFPictureData pictureData = picture.getPictureData();
    byte[] bytes = pictureData.getData();
    ExcelPictureFormat format = ExcelPictureFormat.fromPoiPictureType(pictureData.getPictureType());
    return new ExcelDrawingObjectPayload.Picture(
        objectName,
        format,
        pictureData.getPackagePart().getContentType(),
        objectName + format.defaultExtension(),
        sha256(bytes),
        new ExcelBinaryData(bytes),
        nullIfBlank(picture.getCTPicture().getNvPicPr().getCNvPr().getDescr()));
  }

  private ExcelDrawingObjectPayload.EmbeddedObject embeddedObjectPayload(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    EmbeddedObjectReadback readback = embeddedObjectReadback(objectName, objectData);
    return new ExcelDrawingObjectPayload.EmbeddedObject(
        objectName,
        readback.packagingKind(),
        readback.contentType(),
        readback.fileName(),
        sha256(readback.payload().bytes()),
        readback.payload(),
        readback.label(),
        readback.command());
  }

  private EmbeddedObjectReadback embeddedObjectReadback(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    org.apache.poi.openxml4j.opc.PackagePart previewPart = previewImagePart(objectData);
    ExcelBinaryData previewImage =
        previewPart == null ? null : binary(partBytes(previewPart), "preview image");
    ExcelPictureFormat previewFormat =
        previewPart == null
            ? null
            : ExcelPictureFormat.fromContentType(previewPart.getContentType());
    org.apache.poi.openxml4j.opc.PackagePart objectPart = oleObjectPart(objectData);
    if (objectPart == null) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its package relationship");
    }
    ExcelBinaryData rawPackage;
    rawPackage = binary(partBytes(objectPart), "embedded object package");
    String contentType = objectPart.getContentType();
    String label = null;
    String fileName = partFileName(objectPart);
    String command = null;
    ExcelEmbeddedObjectPackagingKind packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    ExcelBinaryData payload = rawPackage;
    try {
      if (looksLikeOle2Storage(rawPackage.bytes())) {
        org.apache.poi.poifs.filesystem.Ole10Native nativeData = ole10Native(rawPackage.bytes());
        payload = new ExcelBinaryData(nativeData.getDataBuffer());
        label = firstNonBlank(nativeData.getLabel2(), nativeData.getLabel());
        fileName = firstNonBlank(nativeData.getFileName2(), nativeData.getFileName());
        command = firstNonBlank(nativeData.getCommand2(), nativeData.getCommand());
        packagingKind = ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE;
        contentType = "application/octet-stream";
      }
    } catch (IOException | org.apache.poi.poifs.filesystem.Ole10NativeException exception) {
      // Preserve truthful raw-package readback when the embedded object is not a POI-packaged
      // org.apache.poi.poifs.filesystem.Ole10Native payload.
      payload = rawPackage;
      packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    }
    fileName = Objects.requireNonNullElse(fileName, objectName + ".bin");
    return new EmbeddedObjectReadback(
        packagingKind, label, fileName, command, contentType, payload, previewFormat, previewImage);
  }

  private void deleteNamedShapeIfPresent(XSSFSheet sheet, String objectName) {
    LocatedShape located = optionalLocatedShape(sheet, objectName);
    if (located != null) {
      deleteLocatedShape(sheet, located);
    }
  }

  private void deleteLocatedShape(XSSFSheet sheet, LocatedShape located) {
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture picture) {
      deletePicture(sheet, located, picture);
      return;
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      XSSFChart chart = chartForGraphicFrame(located.drawing(), graphicFrame);
      if (chart != null) {
        deleteChart(sheet, located, chart);
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
      cleanupEmptyDrawingPatriarch(sheet);
      return;
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + resolvedName(located.shape())
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is read-only until a later parity phase");
  }

  private void deleteChart(XSSFSheet sheet, LocatedShape located, XSSFChart chart) {
    if (!removePoiRelation(located.drawing(), chart)) {
      throw new IllegalStateException(
          "Failed to remove chart relation for '"
              + resolvedName((org.apache.poi.xssf.usermodel.XSSFGraphicFrame) located.shape())
              + "'");
    }
    removeParentAnchor(located.drawing(), located.parentAnchor());
    cleanupEmptyDrawingPatriarch(sheet);
  }

  private void deletePicture(XSSFSheet sheet, LocatedShape located, XSSFPicture picture) {
    org.apache.poi.openxml4j.opc.PackagePartName imagePartName =
        picture.getPictureData().getPackagePart().getPartName();
    String relationId = picture.getCTPicture().getBlipFill().getBlip().getEmbed();
    located.drawing().getPackagePart().removeRelationship(relationId);
    removeParentAnchor(located.drawing(), located.parentAnchor());
    cleanupWorkbookImagePartIfUnused(sheet.getWorkbook(), imagePartName);
    cleanupEmptyDrawingPatriarch(sheet);
  }

  private void deleteEmbeddedObject(
      XSSFSheet sheet,
      LocatedShape located,
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    org.apache.poi.openxml4j.opc.PackagePart objectPart = oleObjectPart(objectData);
    org.apache.poi.openxml4j.opc.PackagePart previewPart = previewImagePart(objectData);
    org.apache.poi.openxml4j.opc.PackagePartName olePartName =
        objectPart == null ? null : objectPart.getPartName();
    org.apache.poi.openxml4j.opc.PackagePartName previewPartName =
        previewPart == null ? null : previewPart.getPartName();
    String drawingPreviewRelationId = previewDrawingRelationId(objectData);
    String sheetPreviewRelationId = previewSheetRelationId(objectData.getOleObject());
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
    cleanupPackagePartIfUnused(sheet.getWorkbook().getPackage(), olePartName);
    cleanupWorkbookImagePartIfUnused(sheet.getWorkbook(), previewPartName);
    cleanupEmptyDrawingPatriarch(sheet);
  }

  private org.apache.poi.openxml4j.opc.PackagePart oleObjectPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return relatedInternalPart(sheetPart(objectData), objectData.getOleObject().getId());
  }

  private org.apache.poi.openxml4j.opc.PackagePart previewImagePart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    String sheetRelationId = previewSheetRelationId(objectData.getOleObject());
    if (sheetRelationId != null) {
      org.apache.poi.openxml4j.opc.PackagePart previewPart =
          relatedInternalPart(sheetPart(objectData), sheetRelationId);
      if (previewPart != null) {
        return previewPart;
      }
    }
    String drawingRelationId = previewDrawingRelationId(objectData);
    return drawingRelationId == null
        ? null
        : relatedInternalPart(objectData.getDrawing().getPackagePart(), drawingRelationId);
  }

  private void removeOleObject(
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

  private void cleanupWorkbookImagePartIfUnused(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    if (imagePartName == null || imagePartUsed(workbook, imagePartName)) {
      return;
    }
    removeRelationshipsToPart(workbook.getPackagePart(), imagePartName);
    cleanupPackagePartIfUnused(workbook.getPackage(), imagePartName);
  }

  private boolean imagePartUsed(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      XSSFDrawing drawing = workbook.getSheetAt(sheetIndex).getDrawingPatriarch();
      if (drawing == null) {
        continue;
      }
      for (XSSFShape shape : drawing.getShapes()) {
        if (shape instanceof XSSFPicture picture
            && picture.getPictureData().getPackagePart().getPartName().equals(imagePartName)) {
          return true;
        }
        if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
          org.apache.poi.openxml4j.opc.PackagePart previewPart = previewImagePart(objectData);
          if (previewPart != null && previewPart.getPartName().equals(imagePartName)) {
            return true;
          }
        }
      }
    }
    return false;
  }

  private void removeRelationshipsToPart(
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

  private org.apache.poi.openxml4j.opc.PackagePart relatedInternalPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart, String relationshipId) {
    if (relationshipId == null || relationshipId.isBlank()) {
      return null;
    }
    org.apache.poi.openxml4j.opc.PackageRelationship relationship =
        sourcePart.getRelationship(relationshipId);
    if (relationship == null
        || relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
      return null;
    }
    try {
      return sourcePart
          .getPackage()
          .getPart(
              org.apache.poi.openxml4j.opc.PackagingURIHelper.createPartName(
                  org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
                      sourcePart.getPartName().getURI(), relationship.getTargetURI())));
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to resolve related package part", exception);
    }
  }

  private void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg,
      org.apache.poi.openxml4j.opc.PackagePartName partName) {
    if (partName == null) {
      return;
    }
    try {
      for (org.apache.poi.openxml4j.opc.PackagePart part : pkg.getParts()) {
        if (part.isRelationshipPart()) {
          continue;
        }
        for (org.apache.poi.openxml4j.opc.PackageRelationship relationship :
            part.getRelationships()) {
          if (relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
            continue;
          }
          if (partName
              .getURI()
              .equals(
                  org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
                      part.getPartName().getURI(), relationship.getTargetURI()))) {
            return;
          }
        }
      }
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to inspect package relationships", exception);
    }
    if (pkg.containPart(partName)) {
      pkg.deletePartRecursive(partName);
    }
  }

  private void removeParentAnchor(XSSFDrawing drawing, XmlObject parentAnchor) {
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

  private void removeTwoCellAnchor(
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

  private void removeOneCellAnchor(
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

  private void removeAbsoluteAnchor(
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

  private LocatedShape requiredLocatedShape(XSSFSheet sheet, String objectName) {
    LocatedShape located = optionalLocatedShape(sheet, objectName);
    if (located == null) {
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    return located;
  }

  private LocatedShape optionalLocatedShape(XSSFSheet sheet, String objectName) {
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return null;
    }
    XSSFShape matchedShape = null;
    XmlObject matchedShapeXml = null;
    for (XSSFShape shape : drawing.getShapes()) {
      if (!resolvedName(shape).equals(objectName)) {
        continue;
      }
      if (matchedShape != null) {
        throw new IllegalArgumentException(
            "Multiple drawing objects named '"
                + objectName
                + "' exist on sheet '"
                + sheet.getSheetName()
                + "'");
      }
      matchedShape = shape;
      matchedShapeXml = shapeXml(shape);
    }
    return matchedShape == null
        ? null
        : new LocatedShape(drawing, matchedShape, matchedShapeXml, parentAnchor(matchedShapeXml));
  }

  private void createChart(XSSFSheet sheet, PreparedChartDefinition definition) {
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    XSSFChart chart = drawing.createChart(anchor);
    chart.getGraphicFrame().setName(definition.name());
    applyCommonChartState(chart, definition);
    switch (definition) {
      case PreparedBarChart bar -> initializeBarChart(chart, bar);
      case PreparedLineChart line -> initializeLineChart(chart, line);
      case PreparedPieChart pie -> initializePieChart(chart, pie);
    }
  }

  private void mutateChart(
      XSSFSheet sheet, LocatedShape located, XSSFChart chart, PreparedChartDefinition definition) {
    ExcelChartSnapshot snapshot =
        snapshotChart(chart, (org.apache.poi.xssf.usermodel.XSSFGraphicFrame) located.shape());
    if (snapshot instanceof ExcelChartSnapshot.Unsupported unsupported) {
      throw new IllegalArgumentException(
          "Chart '"
              + unsupported.name()
              + "' on sheet '"
              + sheet.getSheetName()
              + "' contains unsupported detail and cannot be mutated authoritatively: "
              + unsupported.detail());
    }

    switch ((ExcelChartSnapshot.Supported) snapshot) {
      case ExcelChartSnapshot.Bar _ -> {
        if (!(definition instanceof PreparedBarChart bar)) {
          throw typeChangeNotSupported(sheet, definition.name());
        }
        updateAnchorInPlace(sheet, definition.name(), located.parentAnchor(), definition.anchor());
        applyCommonChartState(chart, definition);
        mutateBarChart(chart, bar);
      }
      case ExcelChartSnapshot.Line _ -> {
        if (!(definition instanceof PreparedLineChart line)) {
          throw typeChangeNotSupported(sheet, definition.name());
        }
        updateAnchorInPlace(sheet, definition.name(), located.parentAnchor(), definition.anchor());
        applyCommonChartState(chart, definition);
        mutateLineChart(chart, line);
      }
      case ExcelChartSnapshot.Pie _ -> {
        if (!(definition instanceof PreparedPieChart pie)) {
          throw typeChangeNotSupported(sheet, definition.name());
        }
        updateAnchorInPlace(sheet, definition.name(), located.parentAnchor(), definition.anchor());
        applyCommonChartState(chart, definition);
        mutatePieChart(chart, pie);
      }
    }
  }

  private IllegalArgumentException typeChangeNotSupported(XSSFSheet sheet, String chartName) {
    return new IllegalArgumentException(
        "Changing chart type for existing chart '"
            + chartName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is unsupported; delete and recreate the chart to preserve unsupported detail.");
  }

  private void applyCommonChartState(XSSFChart chart, PreparedChartDefinition definition) {
    applyChartTitle(chart, definition.title());
    applyChartLegend(chart, definition.legend());
    chart.displayBlanksAs(ExcelChartPoiBridge.toPoiDisplayBlanks(definition.displayBlanksAs()));
    chart.setPlotOnlyVisibleCells(definition.plotOnlyVisibleCells());
  }

  private void initializeBarChart(XSSFChart chart, PreparedBarChart bar) {
    org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis categoryAxis =
        chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
    org.apache.poi.xddf.usermodel.chart.XDDFValueAxis valueAxis =
        chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
    valueAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);
    org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData)
            chart.createData(
                org.apache.poi.xddf.usermodel.chart.ChartTypes.BAR, categoryAxis, valueAxis);
    data.setBarDirection(ExcelChartPoiBridge.toPoiBarDirection(bar.barDirection()));
    data.setVaryColors(bar.varyColors());
    syncSeries(chart, data, bar.series());
  }

  private void initializeLineChart(XSSFChart chart, PreparedLineChart line) {
    org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis categoryAxis =
        chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
    org.apache.poi.xddf.usermodel.chart.XDDFValueAxis valueAxis =
        chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
    valueAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);
    org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData)
            chart.createData(
                org.apache.poi.xddf.usermodel.chart.ChartTypes.LINE, categoryAxis, valueAxis);
    data.setVaryColors(line.varyColors());
    syncSeries(chart, data, line.series());
  }

  private void initializePieChart(XSSFChart chart, PreparedPieChart pie) {
    org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData)
            chart.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.PIE, null, null);
    data.setVaryColors(pie.varyColors());
    data.setFirstSliceAngle(pie.firstSliceAngle());
    syncSeries(chart, data, pie.series());
  }

  private void mutateBarChart(XSSFChart chart, PreparedBarChart bar) {
    org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData) chart.getChartSeries().getFirst();
    data.setBarDirection(ExcelChartPoiBridge.toPoiBarDirection(bar.barDirection()));
    data.setVaryColors(bar.varyColors());
    syncSeries(chart, data, bar.series());
  }

  private void mutateLineChart(XSSFChart chart, PreparedLineChart line) {
    org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData) chart.getChartSeries().getFirst();
    data.setVaryColors(line.varyColors());
    syncSeries(chart, data, line.series());
  }

  private void mutatePieChart(XSSFChart chart, PreparedPieChart pie) {
    org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData) chart.getChartSeries().getFirst();
    data.setVaryColors(pie.varyColors());
    data.setFirstSliceAngle(pie.firstSliceAngle());
    syncSeries(chart, data, pie.series());
  }

  private void syncSeries(
      XSSFChart chart,
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data,
      List<PreparedChartSeries> definitions) {
    while (data.getSeriesCount() > definitions.size()) {
      data.removeSeries(data.getSeriesCount() - 1);
    }
    for (int index = 0; index < definitions.size(); index++) {
      PreparedChartSeries definition = definitions.get(index);
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series series;
      if (index < data.getSeriesCount()) {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index);
        series.replaceData(definition.categories(), definition.values());
      } else {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series)
                data.addSeries(definition.categories(), definition.values());
      }
      applySeriesTitle(series, definition.title());
    }
    reindexBarSeries(data);
    chart.plot(data);
  }

  private void syncSeries(
      XSSFChart chart,
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data,
      List<PreparedChartSeries> definitions) {
    while (data.getSeriesCount() > definitions.size()) {
      data.removeSeries(data.getSeriesCount() - 1);
    }
    for (int index = 0; index < definitions.size(); index++) {
      PreparedChartSeries definition = definitions.get(index);
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series series;
      if (index < data.getSeriesCount()) {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index);
        series.replaceData(definition.categories(), definition.values());
      } else {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series)
                data.addSeries(definition.categories(), definition.values());
      }
      applySeriesTitle(series, definition.title());
    }
    reindexLineSeries(data);
    chart.plot(data);
  }

  private void syncSeries(
      XSSFChart chart,
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data,
      List<PreparedChartSeries> definitions) {
    while (data.getSeriesCount() > definitions.size()) {
      data.removeSeries(data.getSeriesCount() - 1);
    }
    for (int index = 0; index < definitions.size(); index++) {
      PreparedChartSeries definition = definitions.get(index);
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series series;
      if (index < data.getSeriesCount()) {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index);
        series.replaceData(definition.categories(), definition.values());
      } else {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series)
                data.addSeries(definition.categories(), definition.values());
      }
      applySeriesTitle(series, definition.title());
    }
    reindexPieSeries(data);
    chart.plot(data);
  }

  private PreparedChartDefinition prepareChartDefinition(
      XSSFSheet sheet, ExcelChartDefinition definition) {
    PreparedChartTitle title = prepareChartTitle(sheet, definition.title());
    List<PreparedChartSeries> series = prepareChartSeries(sheet, definition.series());
    return switch (definition) {
      case ExcelChartDefinition.Bar bar ->
          new PreparedBarChart(
              bar.name(),
              bar.anchor(),
              title,
              bar.legend(),
              bar.displayBlanksAs(),
              bar.plotOnlyVisibleCells(),
              bar.varyColors(),
              bar.barDirection(),
              series);
      case ExcelChartDefinition.Line line ->
          new PreparedLineChart(
              line.name(),
              line.anchor(),
              title,
              line.legend(),
              line.displayBlanksAs(),
              line.plotOnlyVisibleCells(),
              line.varyColors(),
              series);
      case ExcelChartDefinition.Pie pie ->
          new PreparedPieChart(
              pie.name(),
              pie.anchor(),
              title,
              pie.legend(),
              pie.displayBlanksAs(),
              pie.plotOnlyVisibleCells(),
              pie.varyColors(),
              pie.firstSliceAngle(),
              series);
    };
  }

  private PreparedChartTitle prepareChartTitle(XSSFSheet sheet, ExcelChartDefinition.Title title) {
    return switch (title) {
      case ExcelChartDefinition.Title.None _ -> new PreparedChartTitleNone();
      case ExcelChartDefinition.Title.Text text -> new PreparedChartTitleText(text.text());
      case ExcelChartDefinition.Title.Formula formula -> {
        CellReference reference =
            resolveSingleCellReference(
                sheet, normalizeFormula(formula.formula()), "Chart title formula");
        yield new PreparedChartTitleFormula(scalarText(sheet, reference), reference);
      }
    };
  }

  private List<PreparedChartSeries> prepareChartSeries(
      XSSFSheet sheet, List<ExcelChartDefinition.Series> definitions) {
    List<PreparedChartSeries> prepared = new ArrayList<>();
    for (ExcelChartDefinition.Series definition : definitions) {
      prepared.add(
          new PreparedChartSeries(
              prepareSeriesTitle(sheet, definition.title()),
              toCategoryDataSource(sheet, definition.categories().formula()),
              toValueDataSource(sheet, definition.values().formula())));
    }
    return List.copyOf(prepared);
  }

  private PreparedSeriesTitle prepareSeriesTitle(
      XSSFSheet sheet, ExcelChartDefinition.Title title) {
    return switch (title) {
      case ExcelChartDefinition.Title.None _ -> new PreparedSeriesTitleNone();
      case ExcelChartDefinition.Title.Text text -> new PreparedSeriesTitleText(text.text());
      case ExcelChartDefinition.Title.Formula formula -> {
        CellReference reference =
            resolveSingleCellReference(
                sheet, normalizeFormula(formula.formula()), "Series title formula");
        yield new PreparedSeriesTitleFormula(scalarText(sheet, reference), reference);
      }
    };
  }

  private org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> toCategoryDataSource(
      XSSFSheet sheet, String formula) {
    ResolvedChartSource source = resolveChartSource(sheet, formula);
    return source.numeric()
        ? org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
            source.numericValues().toArray(Double[]::new), source.referenceFormula())
        : org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
            source.stringValues().toArray(String[]::new), source.referenceFormula());
  }

  private org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<? extends Number>
      toValueDataSource(XSSFSheet sheet, String formula) {
    ResolvedChartSource source = resolveChartSource(sheet, formula);
    if (!source.numeric()) {
      throw new IllegalArgumentException(
          "Chart value source must resolve to numeric cells: " + formula);
    }
    return org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
        source.numericValues().toArray(Double[]::new), source.referenceFormula());
  }

  private ResolvedChartSource resolveChartSource(XSSFSheet sheet, String formula) {
    String normalizedFormula = normalizeFormula(formula);
    ResolvedAreaReference resolved = resolveAreaReference(sheet, normalizedFormula);
    List<String> stringValues = new ArrayList<>();
    List<Double> numericValues = new ArrayList<>();
    boolean numeric = true;
    for (CellReference reference : resolved.areaReference().getAllReferencedCells()) {
      Cell cell =
          resolved.sheet().getRow(reference.getRow()) == null
              ? null
              : resolved.sheet().getRow(reference.getRow()).getCell(reference.getCol());
      CellScalar scalar = scalar(cell);
      if (scalar.kind() == CellScalarKind.NUMERIC) {
        stringValues.add(Double.toString(scalar.number()));
        numericValues.add(scalar.number());
      } else {
        numeric = false;
        stringValues.add(scalar.text());
      }
    }
    return new ResolvedChartSource(
        normalizedFormula,
        resolved.sheet(),
        resolved.areaReference(),
        numeric,
        stringValues,
        numericValues);
  }

  private ResolvedAreaReference resolveAreaReference(XSSFSheet contextSheet, String formula) {
    Name definedName = resolveDefinedNameReference(contextSheet, formula);
    if (definedName != null) {
      return resolveDefinedName(contextSheet, definedName);
    }

    AreaReference[] references =
        parseContiguousAreaReferences(
            normalizeAreaFormulaForPoi(formula),
            "Chart source formula must resolve to one contiguous area: " + formula);
    if (references.length != 1) {
      throw new IllegalArgumentException(
          "Chart source formula must resolve to one contiguous area: " + formula);
    }
    return new ResolvedAreaReference(
        requireAreaReferenceSheet(contextSheet, references[0], formula), references[0]);
  }

  private Name resolveDefinedNameReference(XSSFSheet contextSheet, String formula) {
    if (!formula.matches("^[A-Za-z_][A-Za-z0-9_.]*$")) {
      return null;
    }

    int contextSheetIndex = contextSheet.getWorkbook().getSheetIndex(contextSheet);
    Name workbookScopedMatch = null;
    for (Name name : contextSheet.getWorkbook().getAllNames()) {
      if (!formula.equals(name.getNameName())) {
        continue;
      }
      if (name.getSheetIndex() == contextSheetIndex) {
        return name;
      }
      if (name.getSheetIndex() < 0) {
        workbookScopedMatch = name;
      }
    }
    return workbookScopedMatch;
  }

  private ResolvedAreaReference resolveDefinedName(XSSFSheet contextSheet, Name definedName) {
    String refersToFormula = requiredDefinedNameFormula(definedName);
    AreaReference[] references =
        parseContiguousAreaReferences(
            normalizeAreaFormulaForPoi(refersToFormula),
            "Defined name '" + definedName.getNameName() + "' must resolve to one contiguous area");
    if (references.length != 1) {
      throw new IllegalArgumentException(
          "Defined name '" + definedName.getNameName() + "' must resolve to one contiguous area");
    }
    XSSFSheet resolvedSheet =
        references[0].getFirstCell().getSheetName() == null
            ? requireDefinedNameSheet(contextSheet, definedName)
            : requireSheet(
                contextSheet.getWorkbook(),
                references[0].getFirstCell().getSheetName(),
                definedName.getNameName());
    return new ResolvedAreaReference(resolvedSheet, references[0]);
  }

  private XSSFSheet requireAreaReferenceSheet(
      XSSFSheet contextSheet, AreaReference reference, String formula) {
    String sheetName = reference.getFirstCell().getSheetName();
    return sheetName == null
        ? contextSheet
        : requireSheet(contextSheet.getWorkbook(), sheetName, formula);
  }

  private XSSFSheet requireDefinedNameSheet(XSSFSheet contextSheet, Name definedName) {
    if (definedName.getSheetIndex() >= 0) {
      return contextSheet.getWorkbook().getSheetAt(definedName.getSheetIndex());
    }
    return contextSheet;
  }

  private AreaReference[] parseContiguousAreaReferences(String formula, String failureMessage) {
    try {
      return AreaReference.generateContiguous(SpreadsheetVersion.EXCEL2007, formula);
    } catch (IllegalArgumentException | IllegalStateException exception) {
      throw new IllegalArgumentException(failureMessage, exception);
    }
  }

  private String normalizeAreaFormulaForPoi(String formula) {
    int delimiter = formula.indexOf(':');
    if (delimiter < 0 || delimiter != formula.lastIndexOf(':')) {
      return formula;
    }

    CellReference first = cellReference(formula.substring(0, delimiter));
    CellReference second = cellReference(formula.substring(delimiter + 1));
    if (first == null
        || second == null
        || first.getSheetName() == null
        || !first.getSheetName().equals(second.getSheetName())) {
      return formula;
    }

    return first.formatAsString()
        + ":"
        + new CellReference(
                second.getRow(), second.getCol(), second.isRowAbsolute(), second.isColAbsolute())
            .formatAsString();
  }

  private CellReference cellReference(String value) {
    try {
      return new CellReference(value);
    } catch (IllegalArgumentException exception) {
      return null;
    }
  }

  private XSSFSheet requireSheet(XSSFWorkbook workbook, String sheetName, String formula) {
    XSSFSheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new IllegalArgumentException(
          "Chart source formula '" + formula + "' resolves to missing sheet '" + sheetName + "'");
    }
    return sheet;
  }

  private String normalizeFormula(String formula) {
    String normalized = requireNonBlank(formula, "formula");
    return normalized.startsWith("=") ? normalized.substring(1) : normalized;
  }

  private CellScalar scalar(Cell cell) {
    if (cell == null) {
      return new CellScalar(CellScalarKind.STRING, "", 0d);
    }
    return switch (cell.getCellType()) {
      case STRING -> new CellScalar(CellScalarKind.STRING, cell.getStringCellValue(), 0d);
      case NUMERIC -> new CellScalar(CellScalarKind.NUMERIC, null, cell.getNumericCellValue());
      case BOOLEAN ->
          new CellScalar(CellScalarKind.STRING, Boolean.toString(cell.getBooleanCellValue()), 0d);
      case BLANK, _NONE -> new CellScalar(CellScalarKind.STRING, "", 0d);
      case FORMULA -> scalarFromFormula(cell);
      case ERROR ->
          throw new IllegalArgumentException("Chart source cells must not contain error values");
    };
  }

  CellScalar scalarFromFormula(Cell cell) {
    return switch (cell.getCachedFormulaResultType()) {
      case STRING -> new CellScalar(CellScalarKind.STRING, cell.getStringCellValue(), 0d);
      case NUMERIC -> new CellScalar(CellScalarKind.NUMERIC, null, cell.getNumericCellValue());
      case BOOLEAN ->
          new CellScalar(CellScalarKind.STRING, Boolean.toString(cell.getBooleanCellValue()), 0d);
      case BLANK, _NONE -> new CellScalar(CellScalarKind.STRING, "", 0d);
      case ERROR ->
          throw new IllegalArgumentException("Chart source formulas must not cache error values");
      case FORMULA ->
          throw new IllegalArgumentException(
              "Chart source formulas must expose a cached scalar result");
    };
  }

  static String requiredDefinedNameFormula(Name definedName) {
    Objects.requireNonNull(definedName, "definedName must not be null");
    String refersToFormula = definedName.getRefersToFormula();
    if (refersToFormula == null || refersToFormula.isBlank()) {
      throw new IllegalArgumentException(
          "Defined name '"
              + definedName.getNameName()
              + "' does not resolve to one chart source area");
    }
    return refersToFormula;
  }

  private void applyChartTitle(XSSFChart chart, PreparedChartTitle title) {
    switch (title) {
      case PreparedChartTitleNone _ -> chart.removeTitle();
      case PreparedChartTitleText text -> chart.setTitleText(text.text());
      case PreparedChartTitleFormula formula -> applyChartTitleFormula(chart, formula);
    }
  }

  private void applyChartTitleFormula(XSSFChart chart, PreparedChartTitleFormula formula) {
    Objects.requireNonNull(chart, "chart must not be null");
    Objects.requireNonNull(formula, "formula must not be null");

    org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle title =
        chart.getCTChart().isSetTitle()
            ? chart.getCTChart().getTitle()
            : chart.getCTChart().addNewTitle();
    org.openxmlformats.schemas.drawingml.x2006.chart.CTTx text =
        title.isSetTx() ? title.getTx() : title.addNewTx();
    if (text.isSetRich()) {
      text.unsetRich();
    }

    org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef reference =
        text.isSetStrRef() ? text.getStrRef() : text.addNewStrRef();
    reference.setF(formula.reference().formatAsString());
    writeStringReferenceCache(reference, formula.cachedText());
  }

  private void writeStringReferenceCache(
      org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef reference, String cachedText) {
    Objects.requireNonNull(reference, "reference must not be null");
    Objects.requireNonNull(cachedText, "cachedText must not be null");

    if (reference.isSetStrCache()) {
      reference.unsetStrCache();
    }
    org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData cache = reference.addNewStrCache();
    cache.addNewPtCount().setVal(1);
    cache.addNewPt().setIdx(0);
    cache.getPtArray(0).setV(cachedText);
  }

  private void applyChartLegend(XSSFChart chart, ExcelChartDefinition.Legend legend) {
    switch (legend) {
      case ExcelChartDefinition.Legend.Hidden _ -> chart.deleteLegend();
      case ExcelChartDefinition.Legend.Visible visible ->
          chart
              .getOrAddLegend()
              .setPosition(ExcelChartPoiBridge.toPoiLegendPosition(visible.position()));
    }
  }

  private void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series series,
      PreparedSeriesTitle title) {
    switch (title) {
      case PreparedSeriesTitleNone _ ->
          unsetSeriesTitleIfPresent(series.getCTBarSer()::isSetTx, series.getCTBarSer()::unsetTx);
      case PreparedSeriesTitleText text -> series.setTitle(text.text());
      case PreparedSeriesTitleFormula formula ->
          series.setTitle(formula.cachedText(), formula.reference());
    }
  }

  private void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series series,
      PreparedSeriesTitle title) {
    switch (title) {
      case PreparedSeriesTitleNone _ ->
          unsetSeriesTitleIfPresent(series.getCTLineSer()::isSetTx, series.getCTLineSer()::unsetTx);
      case PreparedSeriesTitleText text -> series.setTitle(text.text());
      case PreparedSeriesTitleFormula formula ->
          series.setTitle(formula.cachedText(), formula.reference());
    }
  }

  private void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series series,
      PreparedSeriesTitle title) {
    switch (title) {
      case PreparedSeriesTitleNone _ ->
          unsetSeriesTitleIfPresent(series.getCTPieSer()::isSetTx, series.getCTPieSer()::unsetTx);
      case PreparedSeriesTitleText text -> series.setTitle(text.text());
      case PreparedSeriesTitleFormula formula ->
          series.setTitle(formula.cachedText(), formula.reference());
    }
  }

  private void unsetSeriesTitleIfPresent(BooleanSupplier hasTitle, Runnable unsetTitle) {
    Objects.requireNonNull(hasTitle, "hasTitle must not be null");
    Objects.requireNonNull(unsetTitle, "unsetTitle must not be null");
    if (hasTitle.getAsBoolean()) {
      unsetTitle.run();
    }
  }

  private CellReference resolveSingleCellReference(
      XSSFSheet sheet, String formula, String detailLabel) {
    ResolvedAreaReference resolved = resolveAreaReference(sheet, formula);
    if (!resolved.areaReference().isSingleCell()) {
      throw new IllegalArgumentException(
          detailLabel + " must resolve to a single cell: " + formula);
    }
    CellReference firstCell = resolved.areaReference().getFirstCell();
    return new CellReference(
        resolved.sheet().getSheetName(), firstCell.getRow(), firstCell.getCol(), true, true);
  }

  private String scalarText(XSSFSheet sheet, CellReference reference) {
    Cell cell =
        sheet.getWorkbook().getSheet(reference.getSheetName()).getRow(reference.getRow()) == null
            ? null
            : sheet
                .getWorkbook()
                .getSheet(reference.getSheetName())
                .getRow(reference.getRow())
                .getCell(reference.getCol());
    CellScalar scalar = scalar(cell);
    return scalar.kind() == CellScalarKind.NUMERIC
        ? Double.toString(scalar.number())
        : scalar.text();
  }

  private XSSFChart chartForGraphicFrame(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    for (XSSFChart chart : drawing.getCharts()) {
      if (chart.getGraphicFrame().getId() == graphicFrame.getId()) {
        return chart;
      }
    }
    return null;
  }

  private String resolvedChartName(org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    String name = nullIfBlank(graphicFrame.getName());
    return name != null ? name : "Chart-" + graphicFrame.getId();
  }

  private ExcelChartSnapshot.Title snapshotTitle(XSSFChart chart) {
    if (!chart.getCTChart().isSetTitle()) {
      return new ExcelChartSnapshot.Title.None();
    }
    String formula = chart.getTitleFormula();
    if (formula != null) {
      return new ExcelChartSnapshot.Title.Formula(formula, cachedTitleText(chart, formula));
    }
    String text = chart.getTitleText().getString();
    return text.isBlank()
        ? new ExcelChartSnapshot.Title.None()
        : new ExcelChartSnapshot.Title.Text(text);
  }

  private String cachedTitleText(XSSFChart chart, String formula) {
    if (!chart.getCTChart().isSetTitle()
        || !chart.getCTChart().getTitle().isSetTx()
        || !chart.getCTChart().getTitle().getTx().isSetStrRef()
        || !chart.getCTChart().getTitle().getTx().getStrRef().isSetStrCache()
        || chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().sizeOfPtArray() == 0) {
      return resolvedTitleFormulaText(chart, formula);
    }
    String cachedText =
        chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV();
    return cachedText.isBlank() ? resolvedTitleFormulaText(chart, formula) : cachedText;
  }

  private String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    try {
      XSSFSheet sheet = chart.getGraphicFrame().getDrawing().getSheet();
      return scalarText(sheet, resolveSingleCellReference(sheet, formula, "Chart title formula"));
    } catch (RuntimeException exception) {
      return "";
    }
  }

  private static RasterDimensions rasterDimensions(byte[] bytes) {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      BufferedImage image = ImageIO.read(input);
      return image == null
          ? RasterDimensions.none()
          : new RasterDimensions(image.getWidth(), image.getHeight());
    } catch (IOException exception) {
      return RasterDimensions.none();
    }
  }

  private ExcelChartSnapshot.Legend snapshotLegend(XSSFChart chart) {
    if (!chart.getCTChart().isSetLegend()) {
      return new ExcelChartSnapshot.Legend.Hidden();
    }
    return new ExcelChartSnapshot.Legend.Visible(
        ExcelChartPoiBridge.fromPoiLegendPosition(
            new org.apache.poi.xddf.usermodel.chart.XDDFChartLegend(chart.getCTChart())
                .getPosition()));
  }

  private ExcelChartDisplayBlanksAs snapshotDisplayBlanks(XSSFChart chart) {
    return chart.getCTChart().isSetDispBlanksAs()
        ? ExcelChartPoiBridge.fromPoiDisplayBlanks(chart.getCTChart().getDispBlanksAs().getVal())
        : ExcelChartDisplayBlanksAs.GAP;
  }

  private List<ExcelChartSnapshot.Axis> snapshotAxes(XSSFChart chart) {
    List<ExcelChartSnapshot.Axis> axes = new ArrayList<>();
    for (org.apache.poi.xddf.usermodel.chart.XDDFChartAxis axis : chart.getAxes()) {
      axes.add(
          new ExcelChartSnapshot.Axis(
              ExcelChartPoiBridge.axisKind(axis),
              ExcelChartPoiBridge.fromPoiAxisPosition(axis.getPosition()),
              ExcelChartPoiBridge.fromPoiAxisCrosses(axis.getCrosses()),
              axis.isVisible()));
    }
    return List.copyOf(axes);
  }

  private List<ExcelChartSnapshot.Series> snapshotSeries(
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series value = data.getSeries(index);
      series.add(
          snapshotSeries(
              value,
              ((org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) value)
                  .getCTBarSer()
                  .getTx()));
    }
    return List.copyOf(series);
  }

  private List<ExcelChartSnapshot.Series> snapshotSeries(
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series value = data.getSeries(index);
      series.add(
          snapshotSeries(
              value,
              ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) value)
                  .getCTLineSer()
                  .getTx()));
    }
    return List.copyOf(series);
  }

  private List<ExcelChartSnapshot.Series> snapshotSeries(
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series value = data.getSeries(index);
      series.add(
          snapshotSeries(
              value,
              ((org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) value)
                  .getCTPieSer()
                  .getTx()));
    }
    return List.copyOf(series);
  }

  private ExcelChartSnapshot.Series snapshotSeries(
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series series,
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title) {
    return new ExcelChartSnapshot.Series(
        snapshotSeriesTitle(title),
        snapshotDataSource(series.getCategoryData()),
        snapshotDataSource(series.getValuesData()));
  }

  private ExcelChartSnapshot.Title snapshotSeriesTitle(
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title) {
    if (title == null) {
      return new ExcelChartSnapshot.Title.None();
    }
    if (title.isSetStrRef()) {
      String cachedText =
          title.getStrRef().isSetStrCache() && title.getStrRef().getStrCache().sizeOfPtArray() > 0
              ? title.getStrRef().getStrCache().getPtArray(0).getV()
              : "";
      return new ExcelChartSnapshot.Title.Formula(title.getStrRef().getF(), cachedText);
    }
    return title.isSetV()
        ? new ExcelChartSnapshot.Title.Text(title.getV())
        : new ExcelChartSnapshot.Title.None();
  }

  private ExcelChartSnapshot.DataSource snapshotDataSource(
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    if (source == null) {
      throw new IllegalStateException("Chart series is missing its data source");
    }
    List<String> values = new ArrayList<>();
    for (int index = 0; index < source.getPointCount(); index++) {
      Object point;
      try {
        point = source.getPointAt(index);
      } catch (IndexOutOfBoundsException exception) {
        point = null;
      }
      values.add(point == null ? "" : point.toString());
    }
    if (source.isReference()) {
      return source.isNumeric()
          ? new ExcelChartSnapshot.DataSource.NumericReference(
              source.getDataRangeReference(), source.getFormatCode(), values)
          : new ExcelChartSnapshot.DataSource.StringReference(
              source.getDataRangeReference(), values);
    }
    return source.isNumeric()
        ? new ExcelChartSnapshot.DataSource.NumericLiteral(source.getFormatCode(), values)
        : new ExcelChartSnapshot.DataSource.StringLiteral(values);
  }

  private List<String> chartPlotTypeTokens(XSSFChart chart) {
    return chartPlotTypeTokens(chart.getChartSeries());
  }

  private List<String> chartPlotTypeTokens(
      List<org.apache.poi.xddf.usermodel.chart.XDDFChartData> chartData) {
    List<String> tokens = new ArrayList<>();
    for (org.apache.poi.xddf.usermodel.chart.XDDFChartData value : chartData) {
      tokens.add(ExcelChartPoiBridge.plotTypeToken(value));
    }
    return List.copyOf(tokens);
  }

  private String titleSummary(ExcelChartSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelChartSnapshot.Bar bar -> titleSummary(bar.title());
      case ExcelChartSnapshot.Line line -> titleSummary(line.title());
      case ExcelChartSnapshot.Pie pie -> titleSummary(pie.title());
      case ExcelChartSnapshot.Unsupported _ -> "";
    };
  }

  private String titleSummary(ExcelChartSnapshot.Title title) {
    return switch (title) {
      case ExcelChartSnapshot.Title.None _ -> "";
      case ExcelChartSnapshot.Title.Text text -> text.text();
      case ExcelChartSnapshot.Title.Formula formula ->
          formula.cachedText().isEmpty() ? formula.formula() : formula.cachedText();
    };
  }

  private void reindexBarSeries(org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data) {
    for (int index = 0; index < data.getSeriesCount(); index++) {
      ((org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index))
          .getCTBarSer()
          .getIdx()
          .setVal(index);
      ((org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index))
          .getCTBarSer()
          .getOrder()
          .setVal(index);
    }
  }

  private void reindexLineSeries(org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data) {
    for (int index = 0; index < data.getSeriesCount(); index++) {
      ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index))
          .getCTLineSer()
          .getIdx()
          .setVal(index);
      ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index))
          .getCTLineSer()
          .getOrder()
          .setVal(index);
    }
  }

  private void reindexPieSeries(org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data) {
    for (int index = 0; index < data.getSeriesCount(); index++) {
      ((org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index))
          .getCTPieSer()
          .getIdx()
          .setVal(index);
      ((org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index))
          .getCTPieSer()
          .getOrder()
          .setVal(index);
    }
  }

  private boolean barVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfBarChartArray() > 0
        && chart.getCTChart().getPlotArea().getBarChartArray(0).isSetVaryColors()
        && chart.getCTChart().getPlotArea().getBarChartArray(0).getVaryColors().getVal();
  }

  private boolean lineVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfLineChartArray() > 0
        && chart.getCTChart().getPlotArea().getLineChartArray(0).isSetVaryColors()
        && chart.getCTChart().getPlotArea().getLineChartArray(0).getVaryColors().getVal();
  }

  boolean pieVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfPieChartArray() > 0
        && chart.getCTChart().getPlotArea().getPieChartArray(0).isSetVaryColors()
        && chart.getCTChart().getPlotArea().getPieChartArray(0).getVaryColors().getVal();
  }

  private record ResolvedAreaReference(XSSFSheet sheet, AreaReference areaReference) {
    private ResolvedAreaReference {
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(areaReference, "areaReference must not be null");
    }
  }

  private record ResolvedChartSource(
      String referenceFormula,
      XSSFSheet sheet,
      AreaReference areaReference,
      boolean numeric,
      List<String> stringValues,
      List<Double> numericValues) {
    private ResolvedChartSource {
      Objects.requireNonNull(referenceFormula, "referenceFormula must not be null");
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(areaReference, "areaReference must not be null");
      stringValues =
          List.copyOf(Objects.requireNonNull(stringValues, "stringValues must not be null"));
      numericValues =
          List.copyOf(Objects.requireNonNull(numericValues, "numericValues must not be null"));
    }
  }

  /** Prepared chart payload validated fully before any chart mutation starts. */
  private sealed interface PreparedChartDefinition
      permits PreparedBarChart, PreparedLineChart, PreparedPieChart {
    /** Returns the target chart name. */
    String name();

    /** Returns the target chart anchor. */
    ExcelDrawingAnchor.TwoCell anchor();

    /** Returns the validated chart title payload. */
    PreparedChartTitle title();

    /** Returns the validated legend payload. */
    ExcelChartDefinition.Legend legend();

    /** Returns the validated blank-cell display behavior. */
    ExcelChartDisplayBlanksAs displayBlanksAs();

    /** Returns whether hidden cells stay excluded from the chart. */
    boolean plotOnlyVisibleCells();
  }

  /** Prepared chart-title payload validated before chart creation or mutation. */
  private sealed interface PreparedChartTitle
      permits PreparedChartTitleNone, PreparedChartTitleText, PreparedChartTitleFormula {}

  /** Prepared series-title payload validated before chart creation or mutation. */
  private sealed interface PreparedSeriesTitle
      permits PreparedSeriesTitleNone, PreparedSeriesTitleText, PreparedSeriesTitleFormula {}

  private record PreparedBarChart(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      PreparedChartTitle title,
      ExcelChartDefinition.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      List<PreparedChartSeries> series)
      implements PreparedChartDefinition {
    private PreparedBarChart {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      series = List.copyOf(Objects.requireNonNull(series, "series must not be null"));
    }
  }

  private record PreparedLineChart(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      PreparedChartTitle title,
      ExcelChartDefinition.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      List<PreparedChartSeries> series)
      implements PreparedChartDefinition {
    private PreparedLineChart {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      series = List.copyOf(Objects.requireNonNull(series, "series must not be null"));
    }
  }

  private record PreparedPieChart(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      PreparedChartTitle title,
      ExcelChartDefinition.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      Integer firstSliceAngle,
      List<PreparedChartSeries> series)
      implements PreparedChartDefinition {
    private PreparedPieChart {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      series = List.copyOf(Objects.requireNonNull(series, "series must not be null"));
    }
  }

  private record PreparedChartTitleNone() implements PreparedChartTitle {}

  private record PreparedChartTitleText(String text) implements PreparedChartTitle {
    private PreparedChartTitleText {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  private record PreparedChartTitleFormula(String cachedText, CellReference reference)
      implements PreparedChartTitle {
    private PreparedChartTitleFormula {
      Objects.requireNonNull(cachedText, "cachedText must not be null");
      Objects.requireNonNull(reference, "reference must not be null");
    }
  }

  private record PreparedSeriesTitleNone() implements PreparedSeriesTitle {}

  private record PreparedSeriesTitleText(String text) implements PreparedSeriesTitle {
    private PreparedSeriesTitleText {
      Objects.requireNonNull(text, "text must not be null");
    }
  }

  private record PreparedSeriesTitleFormula(String cachedText, CellReference reference)
      implements PreparedSeriesTitle {
    private PreparedSeriesTitleFormula {
      Objects.requireNonNull(cachedText, "cachedText must not be null");
      Objects.requireNonNull(reference, "reference must not be null");
    }
  }

  private record PreparedChartSeries(
      PreparedSeriesTitle title,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> categories,
      org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<? extends Number> values) {
    private PreparedChartSeries {
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(categories, "categories must not be null");
      Objects.requireNonNull(values, "values must not be null");
    }
  }

  /** Supported scalar kinds extracted from chart source cells. */
  enum CellScalarKind {
    STRING,
    NUMERIC
  }

  record CellScalar(CellScalarKind kind, String text, double number) {
    CellScalar {
      Objects.requireNonNull(kind, "kind must not be null");
      if (kind == CellScalarKind.STRING) {
        Objects.requireNonNull(text, "text must not be null");
      }
    }
  }

  private org.apache.poi.xssf.usermodel.XSSFClientAnchor toPoiAnchor(
      XSSFDrawing drawing, ExcelDrawingAnchor.TwoCell anchor) {
    org.apache.poi.xssf.usermodel.XSSFClientAnchor poiAnchor =
        drawing.createAnchor(
            anchor.from().dx(),
            anchor.from().dy(),
            anchor.to().dx(),
            anchor.to().dy(),
            anchor.from().columnIndex(),
            anchor.from().rowIndex(),
            anchor.to().columnIndex(),
            anchor.to().rowIndex());
    poiAnchor.setAnchorType(toPoiBehavior(anchor.behavior()));
    return poiAnchor;
  }

  private ClientAnchor.AnchorType toPoiBehavior(ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE -> ClientAnchor.AnchorType.MOVE_AND_RESIZE;
      case MOVE_DONT_RESIZE -> ClientAnchor.AnchorType.MOVE_DONT_RESIZE;
      case DONT_MOVE_AND_RESIZE -> ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;
    };
  }

  private ExcelDrawingAnchor snapshotAnchor(XmlObject shapeXml) {
    XmlObject parentAnchor = parentAnchor(shapeXml);
    return switch (parentAnchor) {
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor
              twoCellAnchor ->
          new ExcelDrawingAnchor.TwoCell(
              marker(twoCellAnchor.getFrom()),
              marker(twoCellAnchor.getTo()),
              behavior(twoCellAnchor.getEditAs()));
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor
              oneCellAnchor ->
          new ExcelDrawingAnchor.OneCell(
              marker(oneCellAnchor.getFrom()),
              oneCellAnchor.getExt().getCx(),
              oneCellAnchor.getExt().getCy(),
              ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor
              absoluteAnchor ->
          new ExcelDrawingAnchor.Absolute(
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

  private ExcelDrawingAnchorBehavior behavior(
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

  private ExcelDrawingMarker marker(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker marker) {
    return new ExcelDrawingMarker(
        marker.getCol(),
        marker.getRow(),
        Math.toIntExact(((Number) marker.getColOff()).longValue()),
        Math.toIntExact(((Number) marker.getRowOff()).longValue()));
  }

  @SuppressWarnings("PMD.UseTryWithResources")
  private String previewSheetRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject) {
    org.apache.xmlbeans.XmlCursor cursor = oleObject.newCursor();
    try {
      return cursor.toChild(org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML, "objectPr")
          ? cursor.getAttributeText(OLE_PREVIEW_ID)
          : null;
    } finally {
      cursor.close();
    }
  }

  private String previewDrawingRelationId(org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    if (objectData.getCTShape().getSpPr().getBlipFill() == null
        || objectData.getCTShape().getSpPr().getBlipFill().getBlip() == null) {
      return null;
    }
    return objectData.getCTShape().getSpPr().getBlipFill().getBlip().getEmbed();
  }

  private XmlObject shapeXml(XSSFShape shape) {
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
  private XmlObject parentAnchor(XmlObject shapeXml) {
    org.apache.xmlbeans.XmlCursor cursor = shapeXml.newCursor();
    try {
      return cursor.toParent() ? cursor.getObject() : null;
    } finally {
      cursor.close();
    }
  }

  private String resolvedName(XSSFShape shape) {
    return nullIfBlank(shape.getShapeName()) != null ? shape.getShapeName() : defaultName(shape);
  }

  private String defaultName(XSSFShape shape) {
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

  private int shapeType(String presetGeometryToken) {
    STShapeType.Enum preset = STShapeType.Enum.forString(presetGeometryToken);
    if (preset == null) {
      throw new IllegalArgumentException("Unsupported presetGeometryToken: " + presetGeometryToken);
    }
    return preset.intValue();
  }

  private ExcelBinaryData binary(byte[] bytes, String label) {
    try {
      return new ExcelBinaryData(bytes);
    } catch (IllegalArgumentException exception) {
      throw new IllegalStateException("Missing " + label + " bytes", exception);
    }
  }

  private byte[] partBytes(org.apache.poi.openxml4j.opc.PackagePart part) {
    try {
      return part.getInputStream().readAllBytes();
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to read package part bytes", exception);
    }
  }

  private String sha256(byte[] bytes) {
    try {
      MessageDigest digest = MessageDigest.getInstance("SHA-256");
      return HexFormat.of().formatHex(digest.digest(bytes));
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 digest is unavailable", exception);
    }
  }

  private String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private String nullIfBlank(String value) {
    return value == null || value.isBlank() ? null : value;
  }

  private String firstNonBlank(String first, String second) {
    String normalizedFirst = nullIfBlank(first);
    return normalizedFirst != null ? normalizedFirst : nullIfBlank(second);
  }

  private boolean removePoiRelation(POIXMLDocumentPart parent, POIXMLDocumentPart child) {
    return poiRelationRemover.test(parent, child);
  }

  private void updateAnchorInPlace(
      XSSFSheet sheet,
      String objectName,
      XmlObject parentAnchor,
      ExcelDrawingAnchor.TwoCell anchor) {
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

  private void applyMarker(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker target,
      ExcelDrawingMarker source) {
    target.setCol(source.columnIndex());
    target.setRow(source.rowIndex());
    target.setColOff(source.dx());
    target.setRowOff(source.dy());
  }

  private org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.Enum toPoiEditAs(
      ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.TWO_CELL;
      case MOVE_DONT_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.ONE_CELL;
      case DONT_MOVE_AND_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.ABSOLUTE;
    };
  }

  private org.apache.poi.openxml4j.opc.PackagePart sheetPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ((XSSFSheet) objectData.getDrawing().getParent()).getPackagePart();
  }

  private boolean looksLikeOle2Storage(byte[] bytes) throws IOException {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      return org.apache.poi.poifs.filesystem.FileMagic.valueOf(
              org.apache.poi.poifs.filesystem.FileMagic.prepareToCheckMagic(input))
          == org.apache.poi.poifs.filesystem.FileMagic.OLE2;
    }
  }

  private org.apache.poi.poifs.filesystem.Ole10Native ole10Native(byte[] bytes)
      throws IOException, org.apache.poi.poifs.filesystem.Ole10NativeException {
    try (org.apache.poi.poifs.filesystem.POIFSFileSystem filesystem =
        new org.apache.poi.poifs.filesystem.POIFSFileSystem(new ByteArrayInputStream(bytes))) {
      org.apache.poi.poifs.filesystem.DirectoryNode directory = filesystem.getRoot();
      return org.apache.poi.poifs.filesystem.Ole10Native.createFromEmbeddedOleObject(directory);
    }
  }

  private String partFileName(org.apache.poi.openxml4j.opc.PackagePart part) {
    String partName = part.getPartName().getName();
    return partName.substring(partName.lastIndexOf('/') + 1);
  }

  private record LocatedShape(
      XSSFDrawing drawing, XSSFShape shape, XmlObject shapeXml, XmlObject parentAnchor) {}

  private record EmbeddedObjectReadback(
      ExcelEmbeddedObjectPackagingKind packagingKind,
      String label,
      String fileName,
      String command,
      String contentType,
      ExcelBinaryData payload,
      ExcelPictureFormat previewFormat,
      ExcelBinaryData previewImage) {}

  private record RasterDimensions(Integer widthPixels, Integer heightPixels) {
    private static RasterDimensions none() {
      return new RasterDimensions(null, null);
    }
  }
}
