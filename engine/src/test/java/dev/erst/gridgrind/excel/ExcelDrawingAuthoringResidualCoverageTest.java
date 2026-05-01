package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.nio.charset.StandardCharsets;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Drawing residual authoring and orphan-chart replacement coverage. */
class ExcelDrawingAuthoringResidualCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  void drawingControllerExercisesResidualAuthoringAndOrphanChartReplacementBranches()
      throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Month");
      sheet.getRow(0).createCell(1).setCellValue("Amount");
      sheet.createRow(1).createCell(0).setCellValue("Jan");
      sheet.getRow(1).createCell(1).setCellValue(10d);
      sheet.createRow(2).createCell(0).setCellValue("Feb");
      sheet.getRow(2).createCell(1).setCellValue(20d);

      controller.setPicture(
          sheet,
          new ExcelPictureDefinition(
              "PlainPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(0, 0, 2, 2),
              null));
      controller.setShape(
          sheet,
          new ExcelShapeDefinition(
              "TextlessShape",
              ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
              anchor(3, 0, 5, 2),
              "rect",
              null));
      controller.setShape(
          sheet,
          new ExcelShapeDefinition(
              "ConnectorOnly",
              ExcelAuthoredDrawingShapeKind.CONNECTOR,
              anchor(6, 0, 8, 2),
              null,
              null));
      controller.setEmbeddedObject(
          sheet,
          new ExcelEmbeddedObjectDefinition(
              "AnchoredEmbed",
              "Payload",
              "payload.txt",
              "payload.txt",
              new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
              ExcelPictureFormat.PNG,
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              anchor(9, 0, 12, 4)));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFPicture plainPicture =
          assertInstanceOf(
              XSSFPicture.class,
              drawing.getShapes().stream()
                  .filter(shape -> "PlainPicture".equals(shape.getShapeName()))
                  .findFirst()
                  .orElseThrow());
      assertNotNull(plainPicture);

      XSSFSimpleShape textlessShape =
          assertInstanceOf(
              XSSFSimpleShape.class,
              drawing.getShapes().stream()
                  .filter(shape -> "TextlessShape".equals(shape.getShapeName()))
                  .findFirst()
                  .orElseThrow());
      assertNotNull(textlessShape);

      controller.setDrawingObjectAnchor(sheet, "PlainPicture", anchor(0, 3, 2, 5));
      controller.setDrawingObjectAnchor(sheet, "TextlessShape", anchor(3, 3, 5, 5));
      controller.setDrawingObjectAnchor(sheet, "ConnectorOnly", anchor(6, 3, 8, 5));
      controller.setDrawingObjectAnchor(sheet, "AnchoredEmbed", anchor(9, 5, 12, 9));
      assertNotNull(
          invoke(
              controller,
              "parentAnchor",
              Object.class,
              invoke(controller, "shapeXml", Object.class, plainPicture)));
    }
  }
}
