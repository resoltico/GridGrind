package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingBinarySupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;

/** Integration tests for drawing, picture, and embedded-object sheet workflows. */
class ExcelDrawingControllerTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void drawingObjectsSupportReadMoveDeleteAndRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-drawing-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")))
          .setShape(
              new ExcelShapeDefinition.SimpleShape(
                  "OpsShape", anchor(5, 1, 8, 5), "rect", Optional.of("Queue")))
          .setShape(new ExcelShapeDefinition.Connector("OpsConnector", anchor(9, 1, 11, 4)))
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(12, 1, 15, 6)));

      List<ExcelDrawingObjectSnapshot> snapshots = sheet.drawings().drawingObjects();
      assertEquals(
          List.of("OpsPicture", "OpsShape", "OpsConnector", "OpsEmbed"),
          snapshots.stream().map(ExcelDrawingObjectSnapshot::name).toList());
      ExcelDrawingObjectSnapshot.Picture pictureSnapshot =
          assertInstanceOf(ExcelDrawingObjectSnapshot.Picture.class, snapshots.getFirst());
      assertEquals(1, pictureSnapshot.widthPixels());
      assertEquals(1, pictureSnapshot.heightPixels());
      assertEquals(
          ExcelDrawingShapeKind.SIMPLE_SHAPE,
          assertInstanceOf(ExcelDrawingObjectSnapshot.Shape.class, snapshots.get(1)).kind());
      assertEquals(
          ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
          assertInstanceOf(ExcelDrawingObjectSnapshot.EmbeddedObject.class, snapshots.get(3))
              .packagingKind());

      ExcelDrawingObjectPayload.Picture picturePayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.Picture.class,
              sheet.drawings().drawingObjectPayload("OpsPicture"));
      ExcelDrawingObjectPayload.EmbeddedObject embeddedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              sheet.drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(PNG_PIXEL_BYTES, picturePayload.data().bytes());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), embeddedPayload.data().bytes());

      ExcelDrawingAnchor.TwoCell movedAnchor = anchor(6, 2, 10, 7);
      sheet.drawings().setDrawingObjectAnchor("OpsShape", movedAnchor);
      ExcelDrawingObjectSnapshot.Shape movedShape =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Shape.class,
              sheet.drawings().drawingObjects().stream()
                  .filter(snapshot -> "OpsShape".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(movedAnchor, movedShape.anchor());

      sheet.drawings().deleteDrawingObject("OpsConnector");
      assertEquals(3, sheet.drawings().drawingObjects().size());

      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = reopened.getSheet("Ops").getDrawingPatriarch();
      assertNotNull(drawing);
      assertEquals(
          List.of("OpsPicture", "OpsShape", "OpsEmbed"),
          drawing.getShapes().stream().map(XSSFShape::getShapeName).toList());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelDrawingObjectSnapshot.Picture pictureSnapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Picture.class,
              reopened.sheet("Ops").drawings().drawingObjects().stream()
                  .filter(snapshot -> "OpsPicture".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(1, pictureSnapshot.widthPixels());
      assertEquals(1, pictureSnapshot.heightPixels());
    }
  }

  @Test
  void signatureLinesSupportReadMoveDeleteAndRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-signature-line-roundtrip-");
    ExcelDrawingAnchor.TwoCell movedAnchor = anchor(6, 2, 10, 7);

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setSignatureLine(
              new ExcelSignatureLineDefinition(
                  "OpsSignature",
                  anchor(1, 1, 4, 6),
                  false,
                  "Review the numbers before signing.",
                  "Ada Lovelace",
                  "Finance",
                  "ada@example.com",
                  null,
                  "invalid",
                  Optional.of(ExcelPictureFormat.PNG),
                  Optional.of(new ExcelBinaryData(PNG_PIXEL_BYTES))));

      List<ExcelDrawingObjectSnapshot> snapshots = sheet.drawings().drawingObjects();
      assertEquals(
          List.of("OpsSignature"),
          snapshots.stream().map(ExcelDrawingObjectSnapshot::name).toList());
      ExcelDrawingObjectSnapshot.SignatureLine signatureLine =
          assertInstanceOf(ExcelDrawingObjectSnapshot.SignatureLine.class, snapshots.getFirst());
      assertFalse(signatureLine.allowComments());
      assertEquals("Review the numbers before signing.", signatureLine.signingInstructions());
      assertEquals("Ada Lovelace", signatureLine.suggestedSigner());
      assertEquals(ExcelPictureFormat.PNG, signatureLine.previewFormat());
      assertEquals(400, signatureLine.previewWidthPixels());
      assertEquals(150, signatureLine.previewHeightPixels());

      IllegalArgumentException noPayload =
          assertThrows(
              IllegalArgumentException.class,
              () -> sheet.drawings().drawingObjectPayload("OpsSignature"));
      assertTrue(noPayload.getMessage().contains("does not expose a binary payload"));

      sheet.drawings().setDrawingObjectAnchor("OpsSignature", movedAnchor);
      ExcelDrawingObjectSnapshot.SignatureLine movedSignature =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.SignatureLine.class,
              sheet.drawings().drawingObjects().stream()
                  .filter(snapshot -> "OpsSignature".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(movedAnchor, movedSignature.anchor());

      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFVMLDrawing vmlDrawing = reopened.getSheet("Ops").getVMLDrawing(false);
      assertNotNull(vmlDrawing);
      CTShape signatureShape = firstSignatureShape(vmlDrawing);
      assertEquals("OpsSignature", signatureShape.getAlt());
      assertEquals(STTrueFalse.F, signatureShape.getSignaturelineArray(0).getAllowcomments());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = reopened.sheet("Ops");
      ExcelDrawingObjectSnapshot.SignatureLine signatureLine =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.SignatureLine.class,
              sheet.drawings().drawingObjects().stream()
                  .filter(snapshot -> "OpsSignature".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(movedAnchor, signatureLine.anchor());
      assertEquals("Finance", signatureLine.suggestedSigner2());
      assertEquals("ada@example.com", signatureLine.suggestedSignerEmail());

      sheet.drawings().deleteDrawingObject("OpsSignature");
      reopened.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFVMLDrawing vmlDrawing = reopened.getSheet("Ops").getVMLDrawing(false);
      assertNotNull(vmlDrawing);
      assertEquals(0, signatureShapeCount(vmlDrawing));
    }
  }

  @Test
  void formulaBackedNumericChartTitlePersistsExplicitCacheAcrossSaveAndLoad() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-title-cache-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.cells().setCell("B1", ExcelCellValue.number(1d));
      sheet.cells().setCell("A2", ExcelCellValue.text("Jan"));
      sheet.cells().setCell("A3", ExcelCellValue.text("Feb"));
      sheet.cells().setCell("A4", ExcelCellValue.text("Mar"));
      sheet.cells().setCell("B2", ExcelCellValue.number(10d));
      sheet.cells().setCell("B3", ExcelCellValue.number(12d));
      sheet.cells().setCell("B4", ExcelCellValue.number(14d));
      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "OpsChart",
                  anchor(0, 0, 3, 4),
                  new ExcelChartDefinition.Title.Formula("B1"),
                  new ExcelChartDefinition.Legend.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ExcelChartDefinition.Series(
                          new ExcelChartDefinition.Title.None(),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = reopened.getSheet("Ops").getDrawingPatriarch();
      assertNotNull(drawing);
      XSSFChart chart = drawing.getCharts().getFirst();
      assertEquals("Ops!$B$1", chart.getTitleFormula());
      assertTrue(chart.getCTChart().getTitle().getTx().getStrRef().isSetStrCache());
      assertEquals(
          "1.0",
          chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelChartSnapshot chart = reopened.sheet("Ops").drawings().charts().getFirst();
      ExcelChartSnapshot.Title.Formula title =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, chart.title());
      assertEquals("Ops!$B$1", title.formula());
      assertEquals("1.0", title.cachedText());

      ExcelDrawingObjectSnapshot.Chart drawingChart =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Chart.class,
              reopened.sheet("Ops").drawings().drawingObjects().getFirst());
      assertEquals("1.0", drawingChart.title());
    }
  }

  @Test
  void formulaChartTitleUpdateReplacesExistingRichTextAndExistingStringCache() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-title-rewrite-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Chart");
      seedChartData(sheet);
      sheet.cells().setCell("D1", ExcelCellValue.number(1d));
      sheet.cells().setCell("E1", ExcelCellValue.number(2d));
      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "OpsChart",
                  anchor(4, 1, 11, 16),
                  new ExcelChartDefinition.Title.Text("Roadmap"),
                  new ExcelChartDefinition.Legend.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ExcelChartDefinition.Series(
                          new ExcelChartDefinition.Title.None(),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = workbook.sheet("Chart");
      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "OpsChart",
                  anchor(4, 1, 11, 16),
                  new ExcelChartDefinition.Title.Formula("D1"),
                  new ExcelChartDefinition.Legend.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ExcelChartDefinition.Series(
                          new ExcelChartDefinition.Title.None(),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "OpsChart",
                  anchor(4, 1, 11, 16),
                  new ExcelChartDefinition.Title.Formula("E1"),
                  new ExcelChartDefinition.Legend.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ExcelChartDefinition.Series(
                          new ExcelChartDefinition.Title.None(),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = reopened.getSheet("Chart").getDrawingPatriarch();
      assertNotNull(drawing);
      XSSFChart chart = drawing.getCharts().getFirst();
      assertEquals("Chart!$E$1", chart.getTitleFormula());
      assertFalse(chart.getCTChart().getTitle().getTx().isSetRich());
      assertTrue(chart.getCTChart().getTitle().getTx().isSetStrRef());
      assertTrue(chart.getCTChart().getTitle().getTx().getStrRef().isSetStrCache());
      assertEquals(
          1, chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().sizeOfPtArray());
      assertEquals(
          "2.0",
          chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelChartSnapshot chart = reopened.sheet("Chart").drawings().charts().getFirst();
      ExcelChartSnapshot.Title.Formula title =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, chart.title());
      assertEquals("Chart!$E$1", title.formula());
      assertEquals("2.0", title.cachedText());
    }
  }

  @Test
  void chartAuthoringSupportsMultiPlotComboCharts() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-combo-authored-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Chart");
      seedChartData(sheet);
      sheet
          .drawings()
          .setChart(
              new ExcelChartDefinition(
                  "ComboChart",
                  anchor(4, 1, 11, 16),
                  new ExcelChartDefinition.Title.Text("Roadmap"),
                  new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  List.of(
                      new ExcelChartDefinition.Bar(
                          true,
                          ExcelChartBarDirection.COLUMN,
                          ExcelChartBarGrouping.CLUSTERED,
                          Optional.empty(),
                          Optional.empty(),
                          List.of(
                              new ExcelChartDefinition.Axis(
                                  ExcelChartAxisKind.CATEGORY,
                                  ExcelChartAxisPosition.BOTTOM,
                                  ExcelChartAxisCrosses.AUTO_ZERO,
                                  true),
                              new ExcelChartDefinition.Axis(
                                  ExcelChartAxisKind.VALUE,
                                  ExcelChartAxisPosition.LEFT,
                                  ExcelChartAxisCrosses.AUTO_ZERO,
                                  true)),
                          List.of(
                              new ExcelChartDefinition.Series(
                                  new ExcelChartDefinition.Title.Text("Plan"),
                                  ExcelChartTestSupport.ref("A2:A4"),
                                  ExcelChartTestSupport.ref("B2:B4")))),
                      new ExcelChartDefinition.Line(
                          false,
                          ExcelChartGrouping.STANDARD,
                          List.of(
                              new ExcelChartDefinition.Axis(
                                  ExcelChartAxisKind.CATEGORY,
                                  ExcelChartAxisPosition.BOTTOM,
                                  ExcelChartAxisCrosses.AUTO_ZERO,
                                  true),
                              new ExcelChartDefinition.Axis(
                                  ExcelChartAxisKind.VALUE,
                                  ExcelChartAxisPosition.LEFT,
                                  ExcelChartAxisCrosses.AUTO_ZERO,
                                  true)),
                          List.of(
                              new ExcelChartDefinition.Series(
                                  new ExcelChartDefinition.Title.Text("Actual"),
                                  ExcelChartTestSupport.ref("A2:A4"),
                                  ExcelChartTestSupport.ref("C2:C4")))))));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelChartSnapshot comboChart = reopened.sheet("Chart").drawings().charts().getFirst();
      assertEquals("ComboChart", comboChart.name());
      assertEquals(
          List.of("BAR", "LINE"),
          comboChart.plots().stream()
              .map(
                  plot ->
                      switch (plot) {
                        case ExcelChartSnapshot.Bar _ -> "BAR";
                        case ExcelChartSnapshot.Line _ -> "LINE";
                        case ExcelChartSnapshot.Unsupported unsupported ->
                            unsupported.plotTypeToken();
                        default -> throw new AssertionError("Unexpected plot type: " + plot);
                      })
              .toList());
      ExcelChartSnapshot.Bar barPlot =
          assertInstanceOf(ExcelChartSnapshot.Bar.class, comboChart.plots().getFirst());
      ExcelChartSnapshot.Line linePlot =
          assertInstanceOf(ExcelChartSnapshot.Line.class, comboChart.plots().get(1));
      assertEquals(
          "Plan",
          assertInstanceOf(ExcelChartSnapshot.Title.Text.class, barPlot.series().getFirst().title())
              .text());
      assertEquals(
          "Actual",
          assertInstanceOf(
                  ExcelChartSnapshot.Title.Text.class, linePlot.series().getFirst().title())
              .text());
    }
  }

  @Test
  void commentOperationsPreserveRealDrawingObjects() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet commentOnly = workbook.getOrCreateSheet("Comments");
      commentOnly.annotations().setComment("A1", new ExcelComment("Review", "GridGrind", false));

      assertTrue(commentOnly.cells().snapshotCell("A1").metadata().comment().isPresent());

      commentOnly.annotations().clearComment("A1");
      assertFalse(commentOnly.cells().snapshotCell("A1").metadata().comment().isPresent());

      ExcelSheet withDrawing = workbook.getOrCreateSheet("Ops");
      withDrawing
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      List<ExcelDrawingObjectSnapshot> before = withDrawing.drawings().drawingObjects();

      withDrawing.annotations().setComment("A1", new ExcelComment("Review", "GridGrind", false));
      assertEquals(before, withDrawing.drawings().drawingObjects());

      withDrawing.annotations().clearComment("A1");
      assertEquals(before, withDrawing.drawings().drawingObjects());
    }
  }

  @Test
  void clearCommentSupportsReopenedPoiCommentWorkbooks() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-clear-comment-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Lead");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      var anchor = drawing.createAnchor(64, 24, 448, 96, 0, 0, 3, 3);
      XSSFComment comment = drawing.createCellComment(anchor);
      comment.setString(new XSSFRichTextString("Review"));
      comment.setAuthor("GridGrind");
      sheet.getRow(0).getCell(0).setCellComment(comment);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      workbook.sheet("Ops").annotations().clearComment("A1");
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      assertNull(workbook.getSheet("Ops").getRow(0).getCell(0).getCellComment());
    }
  }

  @Test
  void embeddedObjectReadbackFallsBackToDrawingPreviewWhenSheetPreviewMetadataIsMissing()
      throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-preview-gap-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Ops")
          .drawings()
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(1, 1, 4, 6)));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFObjectData objectData =
          workbook.getSheet("Ops").createDrawingPatriarch().getShapes().stream()
              .filter(XSSFObjectData.class::isInstance)
              .map(XSSFObjectData.class::cast)
              .findFirst()
              .orElseThrow();
      try (XmlCursor cursor = objectData.getOleObject().newCursor()) {
        assertTrue(cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjectPayload("OpsEmbed"));

      assertEquals(ExcelPictureFormat.PNG, snapshot.previewFormat());
      assertNotNull(snapshot.previewByteSize());
      assertNotNull(snapshot.previewSha256());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void embeddedObjectReadbackSurvivesMissingPreviewImageMetadata() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-preview-missing-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Ops")
          .drawings()
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(1, 1, 4, 6)));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFObjectData objectData =
          workbook.getSheet("Ops").createDrawingPatriarch().getShapes().stream()
              .filter(XSSFObjectData.class::isInstance)
              .map(XSSFObjectData.class::cast)
              .findFirst()
              .orElseThrow();
      try (XmlCursor cursor = objectData.getOleObject().newCursor()) {
        assertTrue(cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }
      objectData.getCTShape().getSpPr().unsetBlipFill();
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjectPayload("OpsEmbed"));

      assertNull(snapshot.previewFormat());
      assertNull(snapshot.previewByteSize());
      assertNull(snapshot.previewSha256());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void embeddedObjectReadbackFallsBackWhenSheetPreviewIdTargetsSheetDrawing() throws IOException {
    Path workbookPath =
        XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-preview-sheet-drawing-fallback-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Ops")
          .drawings()
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(1, 1, 4, 6)));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFSheet sheet = workbook.getSheet("Ops");
      XSSFObjectData objectData = firstEmbeddedObject(sheet);
      ExcelDrawingBinarySupport.setPreviewSheetRelationId(
          objectData.getOleObject(), sheetDrawingRelationId(sheet));
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjectPayload("OpsEmbed"));

      assertEquals(ExcelPictureFormat.PNG, snapshot.previewFormat());
      assertNotNull(snapshot.previewByteSize());
      assertNotNull(snapshot.previewSha256());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void deleteEmbeddedObjectPreservesSheetDrawingWhenSheetPreviewIdTargetsSheetDrawing()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      sheet
          .drawings()
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(5, 1, 8, 6)));

      XSSFSheet poiSheet = sheet.xssfSheet();
      XSSFObjectData objectData = firstEmbeddedObject(poiSheet);
      ExcelDrawingBinarySupport.setPreviewSheetRelationId(
          objectData.getOleObject(), sheetDrawingRelationId(poiSheet));

      sheet.drawings().deleteDrawingObject("OpsEmbed");

      assertEquals(
          List.of("OpsPicture"),
          sheet.drawings().drawingObjects().stream()
              .map(ExcelDrawingObjectSnapshot::name)
              .toList());
      assertNotNull(poiSheet.getDrawingPatriarch());
      assertEquals(1, poiSheet.getDrawingPatriarch().getShapes().size());
      assertNotNull(sheet.drawings().drawingObjectPayload("OpsPicture"));
    }
  }

  @Test
  void embeddedObjectReadbackSurvivesEmptyPackageBytes() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-empty-package-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Ops")
          .drawings()
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(1, 1, 4, 6)));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFObjectData objectData =
          workbook.getSheet("Ops").createDrawingPatriarch().getShapes().stream()
              .filter(XSSFObjectData.class::isInstance)
              .map(XSSFObjectData.class::cast)
              .findFirst()
              .orElseThrow();
      org.apache.poi.openxml4j.opc.PackagePart objectPart =
          ExcelDrawingBinarySupport.oleObjectPart(objectData).orElseThrow();
      try (var outputStream = objectPart.getOutputStream()) {
        outputStream.write(new byte[0]);
      }
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjectPayload("OpsEmbed"));

      assertEquals(0L, snapshot.byteSize());
      assertEquals(ExcelDrawingBinarySupport.sha256(new byte[0]), snapshot.sha256());
      assertEquals(0, payload.data().size());
      assertArrayEquals(new byte[0], payload.data().bytes());
    }
  }

  @Test
  void chartOperationsSupportAuthoringMutationAndDeletion() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Chart");
      seedChartData(sheet);
      workbook
          .names()
          .setNamedRange(
              new ExcelNamedRangeDefinition(
                  "ChartCategories",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  ExcelNamedRangeTarget.range("Chart", "A2:A4")))
          .setNamedRange(
              new ExcelNamedRangeDefinition(
                  "ChartActual",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  ExcelNamedRangeTarget.range("Chart", "C2:C4")));

      sheet.drawings().setChart(initialChartDefinition(anchor(4, 1, 11, 16)));

      assertInitialChartSnapshot(sheet);
      assertInitialChartDrawingObject(sheet);

      ExcelDrawingAnchor.TwoCell movedAnchor = anchor(6, 2, 12, 18);
      sheet.drawings().setDrawingObjectAnchor("OpsChart", movedAnchor);
      sheet.drawings().setChart(updatedChartDefinition(movedAnchor));

      assertUpdatedChartSnapshot(sheet, movedAnchor);

      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertPersistedChartWorkbook(workbookPath);

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      reopened.sheet("Chart").drawings().deleteDrawingObject("OpsChart");
      reopened.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = reopened.getSheet("Chart").getDrawingPatriarch();
      assertTrue(drawing == null || drawing.getCharts().isEmpty());
    }
  }

  @Test
  void unsupportedChartsStayReadableAndRejectAuthoritativeMutation() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-unsupported-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Chart");
      seedChartData(sheet);

      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      var anchor = drawing.createAnchor(0, 0, 0, 0, 4, 1, 11, 16);
      XSSFChart chart = drawing.createChart(anchor);
      chart.getGraphicFrame().setName("ComboChart");
      chart.getOrAddLegend().setPosition(LegendPosition.TOP_RIGHT);
      XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
      valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
      var categories =
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4"));
      var barValues =
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4"));
      var lineValues =
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("C2:C4"));
      XDDFBarChartData barData =
          (XDDFBarChartData) chart.createData(ChartTypes.BAR, categoryAxis, valueAxis);
      barData.addSeries(categories, barValues).setTitle("Plan", null);
      chart.plot(barData);
      XDDFLineChartData lineData =
          (XDDFLineChartData) chart.createData(ChartTypes.LINE, categoryAxis, valueAxis);
      lineData.addSeries(categories, lineValues).setTitle("Actual", null);
      chart.plot(lineData);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = workbook.sheet("Chart");
      ExcelChartSnapshot comboChart = sheet.drawings().charts().getFirst();
      assertEquals("ComboChart", comboChart.name());
      assertEquals(
          List.of("BAR", "LINE"),
          comboChart.plots().stream()
              .map(
                  plot ->
                      switch (plot) {
                        case ExcelChartSnapshot.Bar _ -> "BAR";
                        case ExcelChartSnapshot.Line _ -> "LINE";
                        case ExcelChartSnapshot.Unsupported unsupported ->
                            unsupported.plotTypeToken();
                        default -> throw new AssertionError("Unexpected plot type: " + plot);
                      })
              .toList());

      ExcelDrawingObjectSnapshot.Chart drawingChart =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Chart.class, sheet.drawings().drawingObjects().getFirst());
      assertTrue(drawingChart.supported());
      assertEquals(List.of("BAR", "LINE"), drawingChart.plotTypeTokens());

      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.barChart(
                  "ComboChart",
                  anchor(4, 1, 11, 16),
                  new ExcelChartDefinition.Title.Text("Roadmap"),
                  new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                  ExcelChartDisplayBlanksAs.SPAN,
                  false,
                  true,
                  ExcelChartBarDirection.COLUMN,
                  List.of(
                      new ExcelChartDefinition.Series(
                          new ExcelChartDefinition.Title.Formula("B1"),
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      ExcelChartSnapshot replaced = sheet.drawings().charts().getFirst();
      assertEquals("ComboChart", replaced.name());
      assertInstanceOf(
          ExcelChartSnapshot.Bar.class,
          ExcelChartTestSupport.singlePlot(replaced, ExcelChartSnapshot.Bar.class));

      sheet.cells().setCell("F1", ExcelCellValue.text("Touch"));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = workbook.getSheet("Chart").getDrawingPatriarch();
      assertNotNull(drawing);
      assertEquals(1, drawing.getCharts().size());
      assertEquals(1, drawing.getCharts().getFirst().getChartSeries().size());
    }
  }

  @Test
  void failedShapeAndChartValidationIsNonMutating() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Chart");
      seedChartData(sheet);
      seedChartNamedRanges(workbook);

      IllegalArgumentException invalidShape =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet
                      .drawings()
                      .setShape(
                          new ExcelShapeDefinition.SimpleShape(
                              "OpsBrokenShape",
                              anchor(1, 1, 3, 3),
                              "invalid-shape",
                              Optional.empty())));
      assertTrue(invalidShape.getMessage().contains("Unsupported presetGeometryToken"));
      assertEquals(List.of(), sheet.drawings().drawingObjects());
      assertEquals(List.of(), sheet.drawings().charts());

      IllegalArgumentException invalidChartCreate =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet
                      .drawings()
                      .setChart(invalidBarChartDefinition("OpsBrokenChart", anchor(4, 1, 11, 16))));
      assertTrue(
          invalidChartCreate
              .getMessage()
              .contains("Chart value source must resolve to numeric cells"));
      assertEquals(List.of(), sheet.drawings().drawingObjects());
      assertEquals(List.of(), sheet.drawings().charts());

      sheet.drawings().setChart(initialChartDefinition(anchor(4, 1, 11, 16)));
      sheet.drawings().setChart(lineChartDefinition(anchor(7, 3, 13, 20)));
      ExcelChartSnapshot typeChanged = sheet.drawings().charts().getFirst();
      List<ExcelDrawingObjectSnapshot> drawingObjectsBeforeFailure =
          sheet.drawings().drawingObjects();
      List<ExcelChartSnapshot> chartsBeforeFailure = sheet.drawings().charts();
      assertEquals("OpsChart", typeChanged.name());
      assertInstanceOf(
          ExcelChartSnapshot.Line.class,
          ExcelChartTestSupport.singlePlot(typeChanged, ExcelChartSnapshot.Line.class));

      IllegalArgumentException invalidChartMutation =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet
                      .drawings()
                      .setChart(invalidBarChartDefinition("OpsChart", anchor(7, 3, 13, 20))));
      assertTrue(
          invalidChartMutation
              .getMessage()
              .contains("Chart value source must resolve to numeric cells"));
      assertEquals(drawingObjectsBeforeFailure, sheet.drawings().drawingObjects());
      assertEquals(chartsBeforeFailure, sheet.drawings().charts());
    }
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow),
        new ExcelDrawingMarker(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static XSSFObjectData firstEmbeddedObject(XSSFSheet sheet) {
    return sheet.createDrawingPatriarch().getShapes().stream()
        .filter(XSSFObjectData.class::isInstance)
        .map(XSSFObjectData.class::cast)
        .findFirst()
        .orElseThrow();
  }

  private static String sheetDrawingRelationId(XSSFSheet sheet) {
    assertTrue(sheet.getCTWorksheet().isSetDrawing());
    return sheet.getCTWorksheet().getDrawing().getId();
  }

  private static void seedChartData(ExcelSheet sheet) {
    sheet.cells().setCell("A1", ExcelCellValue.text("Month"));
    sheet.cells().setCell("B1", ExcelCellValue.text("Plan"));
    sheet.cells().setCell("C1", ExcelCellValue.text("Actual"));
    sheet.cells().setCell("A2", ExcelCellValue.text("Jan"));
    sheet.cells().setCell("B2", ExcelCellValue.number(10d));
    sheet.cells().setCell("C2", ExcelCellValue.number(12d));
    sheet.cells().setCell("A3", ExcelCellValue.text("Feb"));
    sheet.cells().setCell("B3", ExcelCellValue.number(18d));
    sheet.cells().setCell("C3", ExcelCellValue.number(16d));
    sheet.cells().setCell("A4", ExcelCellValue.text("Mar"));
    sheet.cells().setCell("B4", ExcelCellValue.number(15d));
    sheet.cells().setCell("C4", ExcelCellValue.number(21d));
  }

  private static CTShape firstSignatureShape(XSSFVMLDrawing vmlDrawing) {
    try (XmlCursor cursor = vmlDrawing.getDocument().getXml().newCursor()) {
      for (boolean found = cursor.toFirstChild(); found; found = cursor.toNextSibling()) {
        if (cursor.getObject() instanceof CTShape shape && shape.sizeOfSignaturelineArray() > 0) {
          return shape;
        }
      }
    }
    fail("Expected one signature-line VML shape");
    throw new AssertionError("unreachable");
  }

  private static int signatureShapeCount(XSSFVMLDrawing vmlDrawing) {
    int count = 0;
    try (XmlCursor cursor = vmlDrawing.getDocument().getXml().newCursor()) {
      for (boolean found = cursor.toFirstChild(); found; found = cursor.toNextSibling()) {
        if (cursor.getObject() instanceof CTShape shape && shape.sizeOfSignaturelineArray() > 0) {
          count++;
        }
      }
    }
    return count;
  }

  private static void seedChartData(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Month");
    sheet.getRow(0).createCell(1).setCellValue("Plan");
    sheet.getRow(0).createCell(2).setCellValue("Actual");
    sheet.createRow(1).createCell(0).setCellValue("Jan");
    sheet.getRow(1).createCell(1).setCellValue(10d);
    sheet.getRow(1).createCell(2).setCellValue(12d);
    sheet.createRow(2).createCell(0).setCellValue("Feb");
    sheet.getRow(2).createCell(1).setCellValue(18d);
    sheet.getRow(2).createCell(2).setCellValue(16d);
    sheet.createRow(3).createCell(0).setCellValue("Mar");
    sheet.getRow(3).createCell(1).setCellValue(15d);
    sheet.getRow(3).createCell(2).setCellValue(21d);
  }

  private static void seedChartNamedRanges(ExcelWorkbook workbook) {
    workbook
        .names()
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Chart", "A2:A4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartActual",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Chart", "C2:C4")));
  }

  private static void assertInitialChartSnapshot(ExcelSheet sheet) {
    ExcelChartSnapshot initial = sheet.drawings().charts().getFirst();
    ExcelChartSnapshot.Bar bar =
        ExcelChartTestSupport.singlePlot(initial, ExcelChartSnapshot.Bar.class);
    assertEquals("OpsChart", initial.name());
    assertEquals(new ExcelChartSnapshot.Title.Text("Roadmap"), initial.title());
    assertEquals(ExcelChartDisplayBlanksAs.SPAN, initial.displayBlanksAs());
    assertFalse(initial.plotOnlyVisibleCells());
    assertTrue(bar.varyColors());
    assertEquals(ExcelChartBarDirection.COLUMN, bar.barDirection());
    assertEquals(2, bar.axes().size());
    assertEquals(2, bar.series().size());
    assertEquals(
        "A2:A4",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.StringReference.class,
                bar.series().getFirst().categories())
            .formula());
    ExcelChartSnapshot.Series namedRangeSeries = bar.series().get(1);
    ExcelChartSnapshot.Title.Formula namedRangeSeriesTitle =
        assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, namedRangeSeries.title());
    assertTrue(namedRangeSeriesTitle.formula().endsWith("$C$1"));
    assertEquals("Actual", namedRangeSeriesTitle.cachedText());
    assertEquals(
        "ChartCategories",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.StringReference.class, namedRangeSeries.categories())
            .formula());
    assertEquals(
        "ChartActual",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.NumericReference.class, namedRangeSeries.values())
            .formula());
  }

  private static void assertInitialChartDrawingObject(ExcelSheet sheet) {
    ExcelDrawingObjectSnapshot.Chart drawingChart =
        assertInstanceOf(
            ExcelDrawingObjectSnapshot.Chart.class, sheet.drawings().drawingObjects().getFirst());
    assertTrue(drawingChart.supported());
    assertEquals(List.of("BAR"), drawingChart.plotTypeTokens());
    assertEquals("Roadmap", drawingChart.title());
  }

  private static void assertUpdatedChartSnapshot(
      ExcelSheet sheet, ExcelDrawingAnchor.TwoCell movedAnchor) {
    ExcelChartSnapshot updated = sheet.drawings().charts().getFirst();
    ExcelChartSnapshot.Bar bar =
        ExcelChartTestSupport.singlePlot(updated, ExcelChartSnapshot.Bar.class);
    assertEquals(movedAnchor, updated.anchor());
    assertEquals(new ExcelChartSnapshot.Title.Text("Actual focus"), updated.title());
    assertInstanceOf(ExcelChartSnapshot.Legend.Hidden.class, updated.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, updated.displayBlanksAs());
    assertTrue(updated.plotOnlyVisibleCells());
    assertFalse(bar.varyColors());
    assertEquals(ExcelChartBarDirection.BAR, bar.barDirection());
    assertEquals(1, bar.series().size());
    assertEquals(
        "ChartCategories",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.StringReference.class,
                bar.series().getFirst().categories())
            .formula());
    assertEquals(
        "ChartActual",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.NumericReference.class,
                bar.series().getFirst().values())
            .formula());
  }

  private static void assertPersistedChartWorkbook(Path workbookPath) throws IOException {
    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = reopened.getSheet("Chart").getDrawingPatriarch();
      assertNotNull(drawing);
      drawing.getShapes();
      assertEquals(1, drawing.getCharts().size());
      XSSFChart chart = drawing.getCharts().getFirst();
      assertEquals("OpsChart", chart.getGraphicFrame().getName());
      assertEquals("Actual focus", chart.getTitleText().getString());
      assertFalse(chart.getCTChart().isSetLegend());
      XDDFBarChartData data =
          assertInstanceOf(XDDFBarChartData.class, chart.getChartSeries().getFirst());
      assertEquals(BarDirection.BAR, data.getBarDirection());
      assertEquals(1, data.getSeriesCount());
      assertTrue(chart.isPlotOnlyVisibleCells());
      assertEquals(CellType.STRING, reopened.getSheet("Chart").getRow(0).getCell(0).getCellType());
    }
  }

  private static ExcelChartDefinition initialChartDefinition(ExcelDrawingAnchor.TwoCell anchor) {
    return ExcelChartTestSupport.barChart(
        "OpsChart",
        anchor,
        new ExcelChartDefinition.Title.Text("Roadmap"),
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Formula("B1"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4")),
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Formula("C1"),
                ExcelChartTestSupport.ref("ChartCategories"),
                ExcelChartTestSupport.ref("ChartActual"))));
  }

  private static ExcelChartDefinition updatedChartDefinition(ExcelDrawingAnchor.TwoCell anchor) {
    return ExcelChartTestSupport.barChart(
        "OpsChart",
        anchor,
        new ExcelChartDefinition.Title.Text("Actual focus"),
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.ZERO,
        true,
        false,
        ExcelChartBarDirection.BAR,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Formula("C1"),
                ExcelChartTestSupport.ref("ChartCategories"),
                ExcelChartTestSupport.ref("ChartActual"))));
  }

  private static ExcelChartDefinition lineChartDefinition(ExcelDrawingAnchor.TwoCell anchor) {
    return ExcelChartTestSupport.lineChart(
        "OpsChart",
        anchor,
        new ExcelChartDefinition.Title.Text("Line focus"),
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.ZERO,
        true,
        false,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Formula("C1"),
                ExcelChartTestSupport.ref("ChartCategories"),
                ExcelChartTestSupport.ref("ChartActual"))));
  }

  private static ExcelChartDefinition invalidBarChartDefinition(
      String name, ExcelDrawingAnchor.TwoCell anchor) {
    return ExcelChartTestSupport.barChart(
        name,
        anchor,
        new ExcelChartDefinition.Title.Text("Broken"),
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Text("Broken"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("A2:A4"))));
  }
}
