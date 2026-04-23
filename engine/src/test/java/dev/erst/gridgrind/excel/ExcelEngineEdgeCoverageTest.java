package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;

/** Residual edge coverage for value objects and helper seams introduced by the XLSX rebuilds. */
class ExcelEngineEdgeCoverageTest {
  @Test
  void sheetLayoutLimitHelpersRejectNonPositiveDefaultColumnWidth() {
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelSheetLayoutLimits.requireDefaultColumnWidth(0, "defaultColumnWidth"));
    assertDoesNotThrow(
        () ->
            ExcelSheetLayoutLimits.requireDefaultColumnWidth(
                ExcelSheetLayoutLimits.MAX_DEFAULT_COLUMN_WIDTH, "defaultColumnWidth"));
  }

  @Test
  void signatureLineValueObjectsRejectInvalidInputsAcrossAllSurfaces() {
    ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);

    assertDoesNotThrow(
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                " ",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                " ",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                " ",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                " ",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                " ",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                " ",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                " ",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                null,
                5L,
                "hash",
                400,
                150));
    assertEquals(
        0L,
        new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                0L,
                "hash",
                400,
                150)
            .previewByteSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                -1L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                null,
                null,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                " ",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                -1,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                -1));

    assertDoesNotThrow(
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                " ",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                " ",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                " ",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "image/png",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                " ",
                5L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                null,
                5L,
                "hash",
                400,
                150));
    assertEquals(
        0L,
        new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                0L,
                "hash",
                400,
                150)
            .previewByteSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                -1L,
                "hash",
                400,
                150));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
                "Signature",
                anchor,
                "setup",
                true,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                5L,
                " ",
                400,
                150));

    assertDoesNotThrow(
        () ->
            new ExcelSignatureLineDefinition(
                "Signature",
                anchor,
                false,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(new byte[] {1})));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                " ",
                anchor,
                false,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(new byte[] {1})));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "Signature",
                anchor,
                false,
                " ",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(new byte[] {1})));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "Signature",
                anchor,
                false,
                "instructions",
                " ",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(new byte[] {1})));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "Signature",
                anchor,
                false,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                "line1\nline2\nline3\nline4",
                "invalid",
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(new byte[] {1})));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "Signature", anchor, false, null, null, null, null, null, "invalid", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "Signature",
                anchor,
                false,
                null,
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                ExcelPictureFormat.PNG,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "Signature",
                anchor,
                false,
                null,
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                null,
                new ExcelBinaryData(new byte[] {1})));
  }

  @Test
  void customXmlAndChartValueObjectsRejectInvalidInputs() {
    assertDoesNotThrow(
        () -> new ExcelCustomXmlDataBindingSnapshot("binding", true, 7L, "file", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCustomXmlDataBindingSnapshot(" ", true, 7L, "file", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCustomXmlDataBindingSnapshot("binding", true, -1L, "file", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCustomXmlDataBindingSnapshot("binding", true, 7L, " ", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCustomXmlDataBindingSnapshot("binding", true, 7L, "file", -1L));

    assertDoesNotThrow(() -> new ExcelCustomXmlMappingLocator(1L, null));
    assertDoesNotThrow(() -> new ExcelCustomXmlMappingLocator(null, "Map"));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCustomXmlMappingLocator(null, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCustomXmlMappingLocator(0L, "Map"));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCustomXmlMappingLocator(1L, " "));

    ExcelCustomXmlLinkedCellSnapshot linkedCell =
        new ExcelCustomXmlLinkedCellSnapshot("Sheet1", "A1", "/root/value", "string");
    ExcelCustomXmlLinkedTableSnapshot linkedTable =
        new ExcelCustomXmlLinkedTableSnapshot("Sheet1", "Table1", "Table1", "A1:B2", "/root");
    assertDoesNotThrow(
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                "urn:test",
                "XSD",
                "schema.xsd",
                "<schema/>",
                new ExcelCustomXmlDataBindingSnapshot("binding", true, 7L, "file", 0L),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                0L,
                "Map",
                "Root",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                null,
                null,
                null,
                null,
                null,
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                " ",
                "Root",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                null,
                null,
                null,
                null,
                null,
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                "Map",
                " ",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                null,
                null,
                null,
                null,
                null,
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                " ",
                true,
                false,
                true,
                false,
                true,
                " ",
                null,
                null,
                null,
                null,
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                null,
                " ",
                null,
                null,
                null,
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                null,
                null,
                " ",
                null,
                null,
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                null,
                null,
                null,
                " ",
                null,
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelCustomXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                true,
                false,
                true,
                false,
                true,
                null,
                null,
                null,
                null,
                null,
                List.of((ExcelCustomXmlLinkedCellSnapshot) null),
                List.of(linkedTable)));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Text("Plan"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"),
                null,
                ExcelChartMarkerStyle.CIRCLE,
                (short) 1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Text("Plan"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"),
                null,
                ExcelChartMarkerStyle.CIRCLE,
                (short) 10,
                -1L));
    ExcelChartDefinition.Series defaultTitleSeries =
        new ExcelChartDefinition.Series(
            null, ExcelChartTestSupport.ref("A2:A4"), ExcelChartTestSupport.ref("B2:B4"));
    assertInstanceOf(ExcelChartDefinition.Title.None.class, defaultTitleSeries.title());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral("0.0", List.of("1")),
                null,
                ExcelChartMarkerStyle.CIRCLE,
                (short) 1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral("0.0", List.of("1")),
                null,
                ExcelChartMarkerStyle.CIRCLE,
                (short) 10,
                -1L));
    ExcelChartSnapshot.Series compactSeries =
        new ExcelChartSnapshot.Series(
            new ExcelChartSnapshot.Title.None(),
            new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
            new ExcelChartSnapshot.DataSource.NumericLiteral("0.0", List.of("1")));
    assertEquals(null, compactSeries.smooth());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot(
                " ",
                ExcelChartTestSupport.anchor(1, 1, 4, 6),
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                List.of(
                    new ExcelChartSnapshot.Line(
                        false,
                        ExcelChartGrouping.STANDARD,
                        List.of(
                            new ExcelChartSnapshot.Axis(
                                ExcelChartAxisKind.CATEGORY,
                                ExcelChartAxisPosition.BOTTOM,
                                ExcelChartAxisCrosses.AUTO_ZERO,
                                true),
                            new ExcelChartSnapshot.Axis(
                                ExcelChartAxisKind.VALUE,
                                ExcelChartAxisPosition.LEFT,
                                ExcelChartAxisCrosses.MIN,
                                true)),
                        List.of(compactSeries)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot(
                "Chart",
                ExcelChartTestSupport.anchor(1, 1, 4, 6),
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                List.of()));
  }

  @Test
  void chartSourceSupportAndDrawingControllerCoverResidualBranchPaths() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      sheet.createRow(0).createCell(0).setCellValue("Month");
      sheet.getRow(0).createCell(1).setCellValue("Plan");
      sheet.createRow(1).createCell(0).setCellValue("Jan");
      sheet.getRow(1).createCell(1).setCellValue(10d);
      sheet.createRow(2).createCell(0).setCellValue("Feb");
      sheet.getRow(2).createCell(1).setCellValue(12d);

      XDDFDataSource<?> stringLiteral =
          ExcelChartSourceSupport.toCategoryDataSource(
              sheet, new ExcelChartDefinition.DataSource.StringLiteral(List.of("Jan", "Feb")));
      assertFalse(stringLiteral.isNumeric());
      XDDFDataSource<?> numericLiteral =
          ExcelChartSourceSupport.toCategoryDataSource(
              sheet, new ExcelChartDefinition.DataSource.NumericLiteral(List.of(10d, 12d)));
      assertTrue(numericLiteral.isNumeric());
      assertThrows(
          IllegalArgumentException.class,
          () ->
              ExcelChartSourceSupport.toValueDataSource(
                  sheet, new ExcelChartDefinition.DataSource.StringLiteral(List.of("Jan"))));

      Name sheetScoped = workbook.createName();
      sheetScoped.setNameName("SheetScopedValues");
      sheetScoped.setSheetIndex(workbook.getSheetIndex(sheet));
      sheetScoped.setRefersToFormula("B2:B3");
      assertEquals(
          sheet, ExcelChartSourceSupport.resolveAreaReference(sheet, "SheetScopedValues").sheet());

      Name workbookScoped = workbook.createName();
      workbookScoped.setNameName("WorkbookValues");
      workbookScoped.setRefersToFormula("Charts!$B$2:$B$3");
      assertEquals(
          "WorkbookValues",
          ExcelChartSourceSupport.resolveDefinedNameReference(sheet, "WorkbookValues")
              .getNameName());
      assertEquals(
          List.of("10.0", "12.0"),
          ExcelChartSourceSupport.resolveChartSource(sheet, "WorkbookValues").stringValues());

      try (ExcelWorkbook wrappedWorkbook = ExcelWorkbook.wrap(workbook)) {
        ExcelSheet wrappedSheet = wrappedWorkbook.getOrCreateSheet("Ops");
        wrappedSheet.setSignatureLine(
            new ExcelSignatureLineDefinition(
                "Dup",
                ExcelChartTestSupport.anchor(5, 1, 8, 6),
                false,
                "instructions",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "invalid",
                ExcelPictureFormat.PNG,
                new ExcelBinaryData(new byte[] {1})));
        forceSignatureLineName(wrappedWorkbook.sheet("Ops").xssfSheet(), "Dup");
        XSSFDrawing drawing = wrappedWorkbook.sheet("Ops").xssfSheet().createDrawingPatriarch();
        XSSFSimpleShape duplicateShape =
            drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 1, 1, 4, 6));
        duplicateShape.getCTShape().getNvSpPr().getCNvPr().setName("Dup");
        duplicateShape.setText("text");
        ExcelDrawingController controller = new ExcelDrawingController();
        assertTrue(
            new ExcelSignatureLineController()
                .hasNamedSignatureLine(wrappedWorkbook.sheet("Ops").xssfSheet(), "Dup"));
        IllegalArgumentException ambiguousPayload =
            assertThrows(
                IllegalArgumentException.class,
                () ->
                    controller.drawingObjectPayload(
                        wrappedWorkbook.sheet("Ops").xssfSheet(), "Dup"));
        assertTrue(ambiguousPayload.getMessage().contains("Multiple drawing objects named 'Dup'"));
        IllegalArgumentException ambiguousMove =
            assertThrows(
                IllegalArgumentException.class,
                () ->
                    controller.setDrawingObjectAnchor(
                        wrappedWorkbook.sheet("Ops").xssfSheet(),
                        "Dup",
                        ExcelChartTestSupport.anchor(9, 1, 12, 6)));
        assertTrue(ambiguousMove.getMessage().contains("Multiple drawing objects named 'Dup'"));
        IllegalArgumentException ambiguousDelete =
            assertThrows(
                IllegalArgumentException.class,
                () ->
                    controller.deleteDrawingObject(
                        wrappedWorkbook.sheet("Ops").xssfSheet(), "Dup"));
        assertTrue(ambiguousDelete.getMessage().contains("Multiple drawing objects named 'Dup'"));
      }
    }
  }

  private static void forceSignatureLineName(XSSFSheet sheet, String name) throws IOException {
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    assertTrue(vmlDrawing != null, "expected test signature line VML drawing");
    List<CTShape> signatureShapes = signatureShapes(vmlDrawing);
    assertFalse(signatureShapes.isEmpty(), "expected at least one signature shape");
    CTShape shape = signatureShapes.getFirst();
    shape.setAlt(name);
    if (shape.sizeOfImagedataArray() > 0) {
      shape.getImagedataArray(0).setTitle(name);
    }
  }

  private static List<CTShape> signatureShapes(XSSFVMLDrawing vmlDrawing) {
    List<CTShape> shapes = new java.util.ArrayList<>();
    try (XmlCursor cursor = vmlDrawing.getDocument().getXml().newCursor()) {
      for (boolean found = cursor.toFirstChild(); found; found = cursor.toNextSibling()) {
        XmlObject object = cursor.getObject();
        if (object instanceof CTShape shape && shape.sizeOfSignaturelineArray() > 0) {
          shapes.add(shape);
        }
      }
    }
    return List.copyOf(shapes);
  }
}
