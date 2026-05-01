package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Drawing chart-frame lookup coverage. */
class ExcelDrawingChartFrameCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  void drawingControllerChartFrameLookupDistinguishesOrphanedFramesFromLiveCharts()
      throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();
    java.nio.file.Path workbookPath =
        XlsxRoundTrip.newWorkbookPath("gridgrind-drawing-orphan-frame-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      org.apache.poi.xssf.usermodel.XSSFChart orphanChart =
          drawing.createChart(poiAnchor(drawing, 0, 0, 4, 6));
      orphanChart.getGraphicFrame().setName("OrphanFrame");
      org.apache.poi.xssf.usermodel.XSSFChart liveChart =
          drawing.createChart(poiAnchor(drawing, 5, 0, 9, 6));
      liveChart.getGraphicFrame().setName("LiveFrame");
      try (var output = java.nio.file.Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    try (var fileSystem = java.nio.file.FileSystems.newFileSystem(workbookPath)) {
      java.nio.file.Path relationshipsPath =
          fileSystem.getPath("/xl/drawings/_rels/drawing1.xml.rels");
      String relationships = java.nio.file.Files.readString(relationshipsPath);
      String updatedRelationships =
          relationships.replaceFirst("<Relationship[^>]+Type=\"[^\"]+/chart\"[^>]*/>", "");
      java.nio.file.Files.writeString(relationshipsPath, updatedRelationships);
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      XSSFSheet sheet = workbook.xssfWorkbook().getSheet("Charts");
      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame orphanFrame =
          assertInstanceOf(
              XSSFGraphicFrame.class,
              drawing.getShapes().stream()
                  .filter(shape -> "OrphanFrame".equals(shape.getShapeName()))
                  .findFirst()
                  .orElseThrow());
      assertNull(ExcelDrawingChartSupport.chartForGraphicFrame(drawing, orphanFrame));
      ExcelDrawingObjectSnapshot.Shape orphanSnapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Shape.class,
              controller.drawingObjects(sheet).stream()
                  .filter(snapshot -> "OrphanFrame".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(ExcelDrawingShapeKind.GRAPHIC_FRAME, orphanSnapshot.kind());
      IllegalArgumentException moveFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.setDrawingObjectAnchor(sheet, "OrphanFrame", anchor(10, 1, 14, 6)));
      assertTrue(moveFailure.getMessage().contains("read-only"));
      IllegalArgumentException deleteFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteDrawingObject(sheet, "OrphanFrame"));
      assertTrue(deleteFailure.getMessage().contains("read-only"));
    }
  }
}
