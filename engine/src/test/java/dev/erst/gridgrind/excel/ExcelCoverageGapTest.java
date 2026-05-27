package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import java.util.Optional;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused closure tests for residual line and branch gaps left by broader behavior suites. */
class ExcelCoverageGapTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  void chartMarkerSymbolTokensTreatNullBlankAndUnknownValuesAsAbsent() {
    assertEquals(Optional.empty(), ExcelChartMarkerStylePoiBridge.fromSymbolToken(null));
    assertEquals(Optional.empty(), ExcelChartMarkerStylePoiBridge.fromSymbolToken(" "));
    assertEquals(Optional.empty(), ExcelChartMarkerStylePoiBridge.fromSymbolToken("mystery"));
    assertEquals(
        Optional.of(ExcelChartMarkerStyle.TRIANGLE),
        ExcelChartMarkerStylePoiBridge.fromSymbolToken("triangle"));
  }

  @Test
  void clonePreparationRepairsExternalHyperlinksWithDanglingRelationshipIds() throws Exception {
    ExcelSheetClonePreparationSupport support = new ExcelSheetClonePreparationSupport();
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sourceSheet = workbook.createSheet("Source");
      sourceSheet.createRow(1).createCell(0).setCellValue("Link");
      XSSFHyperlink hyperlink =
          workbook
              .getCreationHelper()
              .createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.URL);
      hyperlink.setAddress("https://example.com/repaired");
      sourceSheet.getRow(1).getCell(0).setHyperlink(hyperlink);
      sourceSheet.getHyperlinkList().getFirst().getCTHyperlink().setId("rId404");

      support.prepareSourceSheetForClone(sourceSheet);

      String repairedRelationId =
          sourceSheet.getHyperlinkList().getFirst().getCTHyperlink().getId();
      assertEquals(
          "https://example.com/repaired",
          sourceSheet
              .getPackagePart()
              .getRelationship(repairedRelationId)
              .getTargetURI()
              .toString());
      assertDoesNotThrow(() -> workbook.cloneSheet(workbook.getSheetIndex(sourceSheet), "Replica"));
    }
  }

  @Test
  void pivotControllerWrappersCoverCacheLookupAllocatorAndUnusedPartCleanup() throws Exception {
    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivotTable = workbook.xssfWorkbook().getPivotTables().getFirst();
      assertSame(
          pivotTable.getPivotCacheDefinition(),
          controller.cacheDefinition(pivotTable).orElseThrow());

      controller.primePivotTableAllocator(workbook.xssfWorkbook(), Optional.of(pivotTable));

      workbook
          .xssfWorkbook()
          .getPackagePart()
          .addExternalRelationship(
              "https://example.com/pivot", XSSFRelation.WORKBOOK.getRelation());
      PackagePart orphanPart =
          workbook
              .xssfWorkbook()
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/pivotCache/pivotCacheRecords404.xml"),
                  "application/xml");
      assertTrue(workbook.xssfWorkbook().getPackage().containPart(orphanPart.getPartName()));

      controller.cleanupPackagePartIfUnused(
          workbook.xssfWorkbook().getPackage(), orphanPart.getPartName());

      assertFalse(workbook.xssfWorkbook().getPackage().containPart(orphanPart.getPartName()));
    }
  }

  @Test
  void drawingControllerRequiredLocatedShapeReturnsTheExistingShape() throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Drawings");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFSimpleShape shape =
          drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 1, 1, 4, 5));
      shape.getCTShape().getNvSpPr().getCNvPr().setName("Target");

      ExcelDrawingController.LocatedShape located =
          controller.requiredLocatedShape(sheet, "Target");

      assertSame(drawing, located.drawing());
      assertEquals("Target", located.shape().getShapeName());
    }
  }
}
