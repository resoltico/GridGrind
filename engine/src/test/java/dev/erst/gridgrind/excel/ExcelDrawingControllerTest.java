package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.junit.jupiter.api.Test;

/** Integration tests for drawing, picture, and embedded-object sheet workflows. */
class ExcelDrawingControllerTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void drawingObjectsSupportReadMoveDeleteAndRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-drawing-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  "Queue preview"))
          .setShape(
              new ExcelShapeDefinition(
                  "OpsShape",
                  ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                  anchor(5, 1, 8, 5),
                  "rect",
                  "Queue"))
          .setShape(
              new ExcelShapeDefinition(
                  "OpsConnector",
                  ExcelAuthoredDrawingShapeKind.CONNECTOR,
                  anchor(9, 1, 11, 4),
                  null,
                  null))
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

      List<ExcelDrawingObjectSnapshot> snapshots = sheet.drawingObjects();
      assertEquals(
          List.of("OpsPicture", "OpsShape", "OpsConnector", "OpsEmbed"),
          snapshots.stream().map(ExcelDrawingObjectSnapshot::name).toList());
      assertEquals(
          ExcelDrawingShapeKind.SIMPLE_SHAPE,
          assertInstanceOf(ExcelDrawingObjectSnapshot.Shape.class, snapshots.get(1)).kind());
      assertEquals(
          ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
          assertInstanceOf(ExcelDrawingObjectSnapshot.EmbeddedObject.class, snapshots.get(3))
              .packagingKind());

      ExcelDrawingObjectPayload.Picture picturePayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.Picture.class, sheet.drawingObjectPayload("OpsPicture"));
      ExcelDrawingObjectPayload.EmbeddedObject embeddedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              sheet.drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(PNG_PIXEL_BYTES, picturePayload.data().bytes());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), embeddedPayload.data().bytes());

      ExcelDrawingAnchor.TwoCell movedAnchor = anchor(6, 2, 10, 7);
      sheet.setDrawingObjectAnchor("OpsShape", movedAnchor);
      ExcelDrawingObjectSnapshot.Shape movedShape =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Shape.class,
              sheet.drawingObjects().stream()
                  .filter(snapshot -> "OpsShape".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(movedAnchor, movedShape.anchor());

      sheet.deleteDrawingObject("OpsConnector");
      assertEquals(3, sheet.drawingObjects().size());

      workbook.save(workbookPath);
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = reopened.getSheet("Ops").getDrawingPatriarch();
      assertNotNull(drawing);
      assertEquals(
          List.of("OpsPicture", "OpsShape", "OpsEmbed"),
          drawing.getShapes().stream().map(XSSFShape::getShapeName).toList());
    }
  }

  @Test
  void commentOperationsPreserveRealDrawingObjects() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet commentOnly = workbook.getOrCreateSheet("Comments");
      commentOnly.setComment("A1", new ExcelComment("Review", "GridGrind", false));

      assertTrue(commentOnly.snapshotCell("A1").metadata().comment().isPresent());

      commentOnly.clearComment("A1");
      assertFalse(commentOnly.snapshotCell("A1").metadata().comment().isPresent());

      ExcelSheet withDrawing = workbook.getOrCreateSheet("Ops");
      withDrawing.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
      List<ExcelDrawingObjectSnapshot> before = withDrawing.drawingObjects();

      withDrawing.setComment("A1", new ExcelComment("Review", "GridGrind", false));
      assertEquals(before, withDrawing.drawingObjects());

      withDrawing.clearComment("A1");
      assertEquals(before, withDrawing.drawingObjects());
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

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      workbook.sheet("Ops").clearComment("A1");
      workbook.save(workbookPath);
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      assertNull(workbook.getSheet("Ops").getRow(0).getCell(0).getCellComment());
    }
  }

  @Test
  void embeddedObjectReadbackFallsBackToDrawingPreviewWhenSheetPreviewMetadataIsMissing()
      throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-preview-gap-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook
          .getOrCreateSheet("Ops")
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
      workbook.save(workbookPath);
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

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjectPayload("OpsEmbed"));

      assertEquals(ExcelPictureFormat.PNG, snapshot.previewFormat());
      assertNotNull(snapshot.previewByteSize());
      assertNotNull(snapshot.previewSha256());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void embeddedObjectReadbackSurvivesMissingPreviewImageMetadata() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-preview-missing-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook
          .getOrCreateSheet("Ops")
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
      workbook.save(workbookPath);
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

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjectPayload("OpsEmbed"));

      assertNull(snapshot.previewFormat());
      assertNull(snapshot.previewByteSize());
      assertNull(snapshot.previewSha256());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void chartOperationsSupportAuthoringMutationAndDeletion() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Chart");
      seedChartData(sheet);
      workbook
          .setNamedRange(
              new ExcelNamedRangeDefinition(
                  "ChartCategories",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  new ExcelNamedRangeTarget("Chart", "A2:A4")))
          .setNamedRange(
              new ExcelNamedRangeDefinition(
                  "ChartActual",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  new ExcelNamedRangeTarget("Chart", "C2:C4")));

      sheet.setChart(initialChartDefinition(anchor(4, 1, 11, 16)));

      assertInitialChartSnapshot(sheet);
      assertInitialChartDrawingObject(sheet);

      ExcelDrawingAnchor.TwoCell movedAnchor = anchor(6, 2, 12, 18);
      sheet.setDrawingObjectAnchor("OpsChart", movedAnchor);
      sheet.setChart(updatedChartDefinition(movedAnchor));

      assertUpdatedChartSnapshot(sheet, movedAnchor);

      workbook.save(workbookPath);
    }

    assertPersistedChartWorkbook(workbookPath);

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      reopened.sheet("Chart").deleteDrawingObject("OpsChart");
      reopened.save(workbookPath);
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

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = workbook.sheet("Chart");
      ExcelChartSnapshot.Unsupported unsupported =
          assertInstanceOf(ExcelChartSnapshot.Unsupported.class, sheet.charts().getFirst());
      assertEquals("ComboChart", unsupported.name());
      assertEquals(List.of("BAR", "LINE"), unsupported.plotTypeTokens());

      ExcelDrawingObjectSnapshot.Chart drawingChart =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Chart.class, sheet.drawingObjects().getFirst());
      assertFalse(drawingChart.supported());
      assertEquals(List.of("BAR", "LINE"), drawingChart.plotTypeTokens());

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      new ExcelChartDefinition.Bar(
                          "ComboChart",
                          anchor(4, 1, 11, 16),
                          new ExcelChartDefinition.Title.Text("Roadmap"),
                          new ExcelChartDefinition.Legend.Visible(
                              ExcelChartLegendPosition.TOP_RIGHT),
                          ExcelChartDisplayBlanksAs.SPAN,
                          false,
                          true,
                          ExcelChartBarDirection.COLUMN,
                          List.of(
                              new ExcelChartDefinition.Series(
                                  new ExcelChartDefinition.Title.Formula("B1"),
                                  new ExcelChartDefinition.DataSource("A2:A4"),
                                  new ExcelChartDefinition.DataSource("B2:B4"))))));
      assertTrue(failure.getMessage().contains("unsupported detail"));

      sheet.setCell("F1", ExcelCellValue.text("Touch"));
      workbook.save(workbookPath);
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = workbook.getSheet("Chart").getDrawingPatriarch();
      assertNotNull(drawing);
      assertEquals(1, drawing.getCharts().size());
      assertEquals(2, drawing.getCharts().getFirst().getChartSeries().size());
    }
  }

  @Test
  void failedShapeAndChartValidationIsNonMutating() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Chart");
      seedChartData(sheet);
      seedChartNamedRanges(workbook);

      IllegalArgumentException invalidShape =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setShape(
                      new ExcelShapeDefinition(
                          "OpsBrokenShape",
                          ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                          anchor(1, 1, 3, 3),
                          "invalid-shape",
                          null)));
      assertTrue(invalidShape.getMessage().contains("Unsupported presetGeometryToken"));
      assertEquals(List.of(), sheet.drawingObjects());
      assertEquals(List.of(), sheet.charts());

      IllegalArgumentException invalidChartCreate =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      invalidBarChartDefinition("OpsBrokenChart", anchor(4, 1, 11, 16))));
      assertTrue(
          invalidChartCreate
              .getMessage()
              .contains("Chart value source must resolve to numeric cells"));
      assertEquals(List.of(), sheet.drawingObjects());
      assertEquals(List.of(), sheet.charts());

      sheet.setChart(initialChartDefinition(anchor(4, 1, 11, 16)));
      List<ExcelDrawingObjectSnapshot> drawingObjectsBeforeFailure = sheet.drawingObjects();
      List<ExcelChartSnapshot> chartsBeforeFailure = sheet.charts();

      IllegalArgumentException typeChangeFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> sheet.setChart(lineChartDefinition(anchor(7, 3, 13, 20))));
      assertTrue(typeChangeFailure.getMessage().contains("Changing chart type"));
      assertEquals(drawingObjectsBeforeFailure, sheet.drawingObjects());
      assertEquals(chartsBeforeFailure, sheet.charts());

      IllegalArgumentException invalidChartMutation =
          assertThrows(
              IllegalArgumentException.class,
              () -> sheet.setChart(invalidBarChartDefinition("OpsChart", anchor(7, 3, 13, 20))));
      assertTrue(
          invalidChartMutation
              .getMessage()
              .contains("Chart value source must resolve to numeric cells"));
      assertEquals(drawingObjectsBeforeFailure, sheet.drawingObjects());
      assertEquals(chartsBeforeFailure, sheet.charts());
    }
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow),
        new ExcelDrawingMarker(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static void seedChartData(ExcelSheet sheet) {
    sheet.setCell("A1", ExcelCellValue.text("Month"));
    sheet.setCell("B1", ExcelCellValue.text("Plan"));
    sheet.setCell("C1", ExcelCellValue.text("Actual"));
    sheet.setCell("A2", ExcelCellValue.text("Jan"));
    sheet.setCell("B2", ExcelCellValue.number(10d));
    sheet.setCell("C2", ExcelCellValue.number(12d));
    sheet.setCell("A3", ExcelCellValue.text("Feb"));
    sheet.setCell("B3", ExcelCellValue.number(18d));
    sheet.setCell("C3", ExcelCellValue.number(16d));
    sheet.setCell("A4", ExcelCellValue.text("Mar"));
    sheet.setCell("B4", ExcelCellValue.number(15d));
    sheet.setCell("C4", ExcelCellValue.number(21d));
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
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Chart", "A2:A4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartActual",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Chart", "C2:C4")));
  }

  private static void assertInitialChartSnapshot(ExcelSheet sheet) {
    ExcelChartSnapshot.Bar initial =
        assertInstanceOf(ExcelChartSnapshot.Bar.class, sheet.charts().getFirst());
    assertEquals("OpsChart", initial.name());
    assertEquals(new ExcelChartSnapshot.Title.Text("Roadmap"), initial.title());
    assertEquals(ExcelChartDisplayBlanksAs.SPAN, initial.displayBlanksAs());
    assertFalse(initial.plotOnlyVisibleCells());
    assertTrue(initial.varyColors());
    assertEquals(ExcelChartBarDirection.COLUMN, initial.barDirection());
    assertEquals(2, initial.axes().size());
    assertEquals(2, initial.series().size());
    assertEquals(
        "A2:A4",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.StringReference.class,
                initial.series().getFirst().categories())
            .formula());
    ExcelChartSnapshot.Series namedRangeSeries = initial.series().get(1);
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
        assertInstanceOf(ExcelDrawingObjectSnapshot.Chart.class, sheet.drawingObjects().getFirst());
    assertTrue(drawingChart.supported());
    assertEquals(List.of("BAR"), drawingChart.plotTypeTokens());
    assertEquals("Roadmap", drawingChart.title());
  }

  private static void assertUpdatedChartSnapshot(
      ExcelSheet sheet, ExcelDrawingAnchor.TwoCell movedAnchor) {
    ExcelChartSnapshot.Bar updated =
        assertInstanceOf(ExcelChartSnapshot.Bar.class, sheet.charts().getFirst());
    assertEquals(movedAnchor, updated.anchor());
    assertEquals(new ExcelChartSnapshot.Title.Text("Actual focus"), updated.title());
    assertInstanceOf(ExcelChartSnapshot.Legend.Hidden.class, updated.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, updated.displayBlanksAs());
    assertTrue(updated.plotOnlyVisibleCells());
    assertFalse(updated.varyColors());
    assertEquals(ExcelChartBarDirection.BAR, updated.barDirection());
    assertEquals(1, updated.series().size());
    assertEquals(
        "ChartCategories",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.StringReference.class,
                updated.series().getFirst().categories())
            .formula());
    assertEquals(
        "ChartActual",
        assertInstanceOf(
                ExcelChartSnapshot.DataSource.NumericReference.class,
                updated.series().getFirst().values())
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

  private static ExcelChartDefinition.Bar initialChartDefinition(
      ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelChartDefinition.Bar(
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
                new ExcelChartDefinition.DataSource("A2:A4"),
                new ExcelChartDefinition.DataSource("B2:B4")),
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Formula("C1"),
                new ExcelChartDefinition.DataSource("ChartCategories"),
                new ExcelChartDefinition.DataSource("ChartActual"))));
  }

  private static ExcelChartDefinition.Bar updatedChartDefinition(
      ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelChartDefinition.Bar(
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
                new ExcelChartDefinition.DataSource("ChartCategories"),
                new ExcelChartDefinition.DataSource("ChartActual"))));
  }

  private static ExcelChartDefinition.Line lineChartDefinition(ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelChartDefinition.Line(
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
                new ExcelChartDefinition.DataSource("ChartCategories"),
                new ExcelChartDefinition.DataSource("ChartActual"))));
  }

  private static ExcelChartDefinition.Bar invalidBarChartDefinition(
      String name, ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelChartDefinition.Bar(
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
                new ExcelChartDefinition.DataSource("A2:A4"),
                new ExcelChartDefinition.DataSource("A2:A4"))));
  }
}
