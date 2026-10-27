package dev.erst.gridgrind.engine.backend.xssf;

import static dev.erst.gridgrind.engine.backend.xssf.ExcelSignatureLineSnapshotTestSupport.drawingSignatureLine;
import static dev.erst.gridgrind.engine.backend.xssf.ExcelSignatureLineSnapshotTestSupport.signatureLineSnapshot;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.drawing.SignatureAdapter;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
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
    DrawingAnchorSpec.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);

    assertDoesNotThrow(
        () ->
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
        signatureLineSnapshot(
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
            .preview()
            .orElseThrow()
            .byteSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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
            signatureLineSnapshot(
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

    DrawingObjectReport.SignatureLine drawingSignature =
        assertDoesNotThrow(
            () ->
                drawingSignatureLine(
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
    assertEquals("setup", drawingSignature.setup().orElseThrow().setupId().orElseThrow());
    assertEquals(ExcelPictureFormat.PNG, drawingSignature.preview().orElseThrow().format());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
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
            drawingSignatureLine(
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
        drawingSignatureLine(
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
            .preview()
            .orElseThrow()
            .byteSize());

    assertDoesNotThrow(
        () ->
            new ResolvedSignatureLine(
                "Signature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.of(new BinaryContent(new byte[] {1}))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                " ",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.of(new BinaryContent(new byte[] {1}))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "Signature",
                anchor,
                false,
                Optional.ofNullable(" "),
                Optional.ofNullable("Ada"),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.of(new BinaryContent(new byte[] {1}))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "Signature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable(" "),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.of(new BinaryContent(new byte[] {1}))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "Signature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.ofNullable("line1\nline2\nline3\nline4"),
                Optional.ofNullable("invalid"),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.of(new BinaryContent(new byte[] {1}))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "Signature",
                anchor,
                false,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "Signature",
                anchor,
                false,
                Optional.empty(),
                Optional.ofNullable("Ada"),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "Signature",
                anchor,
                false,
                Optional.empty(),
                Optional.ofNullable("Ada"),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.empty(),
                Optional.of(new BinaryContent(new byte[] {1}))));
  }

  @Test
  void customXmlAndChartValueObjectsRejectInvalidInputs() {
    assertDoesNotThrow(() -> new CustomXmlDataBindingReport("binding", true, 7L, "file", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlDataBindingReport(" ", true, 7L, "file", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlDataBindingReport("binding", true, -1L, "file", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlDataBindingReport("binding", true, 7L, " ", 0L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlDataBindingReport("binding", true, 7L, "file", -1L));

    assertDoesNotThrow(
        () -> new CustomXmlMappingLocator(java.util.Optional.of(1L), java.util.Optional.empty()));
    assertDoesNotThrow(
        () ->
            new CustomXmlMappingLocator(java.util.Optional.empty(), java.util.Optional.of("Map")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlMappingLocator(java.util.Optional.empty(), java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlMappingLocator(java.util.Optional.of(0L), java.util.Optional.of("Map")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlMappingLocator(java.util.Optional.of(1L), java.util.Optional.of(" ")));

    CustomXmlLinkedCellReport linkedCell =
        new CustomXmlLinkedCellReport(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Sheet1"),
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            "/root/value",
            "string");
    CustomXmlLinkedTableReport linkedTable =
        new CustomXmlLinkedTableReport(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Sheet1"),
            "Table1",
            "Table1",
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"),
            "/root");
    assertDoesNotThrow(
        () ->
            customXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                settings(true, false, true, false, true),
                schema(
                    Optional.of("urn:test"),
                    Optional.of("XSD"),
                    Optional.of("schema.xsd"),
                    Optional.of("<schema/>")),
                Optional.of(new CustomXmlDataBindingReport("binding", true, 7L, "file", 0L)),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                0L,
                "Map",
                "Root",
                "Schema",
                settings(true, false, true, false, true),
                schema(Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()),
                Optional.empty(),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                1L,
                " ",
                "Root",
                "Schema",
                settings(true, false, true, false, true),
                schema(Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()),
                Optional.empty(),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                1L,
                "Map",
                " ",
                "Schema",
                settings(true, false, true, false, true),
                schema(Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()),
                Optional.empty(),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                " ",
                settings(true, false, true, false, true),
                schema(Optional.of(" "), Optional.empty(), Optional.empty(), Optional.empty()),
                Optional.empty(),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                settings(true, false, true, false, true),
                schema(Optional.empty(), Optional.of(" "), Optional.empty(), Optional.empty()),
                Optional.empty(),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                settings(true, false, true, false, true),
                schema(Optional.empty(), Optional.empty(), Optional.of(" "), Optional.empty()),
                Optional.empty(),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                settings(true, false, true, false, true),
                schema(Optional.empty(), Optional.empty(), Optional.empty(), Optional.of(" ")),
                Optional.empty(),
                List.of(linkedCell),
                List.of(linkedTable)));
    assertThrows(
        NullPointerException.class,
        () ->
            customXmlMappingSnapshot(
                1L,
                "Map",
                "Root",
                "Schema",
                settings(true, false, true, false, true),
                schema(Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()),
                Optional.empty(),
                List.of((CustomXmlLinkedCellReport) null),
                List.of(linkedTable)));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedChart.Series(
                new ResolvedChartTitle.Text("Plan"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"),
                Optional.empty(),
                Optional.of(ExcelChartMarkerStyle.CIRCLE),
                Optional.of((short) 1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedChart.Series(
                new ResolvedChartTitle.Text("Plan"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"),
                Optional.empty(),
                Optional.of(ExcelChartMarkerStyle.CIRCLE),
                Optional.of((short) 10),
                Optional.of(-1L)));
    ResolvedChart.Series defaultTitleSeries =
        new ResolvedChart.Series(
            null, ExcelChartTestSupport.ref("A2:A4"), ExcelChartTestSupport.ref("B2:B4"));
    assertInstanceOf(ChartTitleSpec.None.class, defaultTitleSeries.title());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.None(),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.of("0.0"), List.of("1")),
                Optional.empty(),
                Optional.of(ExcelChartMarkerStyle.CIRCLE),
                Optional.of((short) 1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.None(),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.of("0.0"), List.of("1")),
                Optional.empty(),
                Optional.of(ExcelChartMarkerStyle.CIRCLE),
                Optional.of((short) 10),
                Optional.of(-1L)));
    ChartReport.Series compactSeries =
        new ChartReport.Series(
            new ChartReport.Title.None(),
            new ChartReport.DataSource.StringLiteral(List.of("Jan")),
            new ChartReport.DataSource.NumericLiteral(Optional.of("0.0"), List.of("1")),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertEquals(Optional.empty(), compactSeries.smooth());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport(
                " ",
                DrawingGeometryFixtures.report(ExcelChartTestSupport.anchor(1, 1, 4, 6)),
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                List.of(
                    new ChartReport.Line(
                        false,
                        ExcelChartGrouping.STANDARD,
                        List.of(
                            new ChartReport.Axis(
                                ExcelChartAxisKind.CATEGORY,
                                ExcelChartAxisPosition.BOTTOM,
                                ExcelChartAxisCrosses.AUTO_ZERO,
                                true),
                            new ChartReport.Axis(
                                ExcelChartAxisKind.VALUE,
                                ExcelChartAxisPosition.LEFT,
                                ExcelChartAxisCrosses.MIN,
                                true)),
                        List.of(compactSeries)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport(
                "Chart",
                DrawingGeometryFixtures.report(ExcelChartTestSupport.anchor(1, 1, 4, 6)),
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
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
          ChartAdapter.toCategoryDataSource(
              sheet, new ChartDataSourceSpec.StringLiteral(List.of("Jan", "Feb")));
      assertFalse(stringLiteral.isNumeric());
      XDDFDataSource<?> numericLiteral =
          ChartAdapter.toCategoryDataSource(
              sheet, new ChartDataSourceSpec.NumericLiteral(List.of(10d, 12d)));
      assertTrue(numericLiteral.isNumeric());
      assertThrows(
          IllegalArgumentException.class,
          () ->
              ChartAdapter.toValueDataSource(
                  sheet, new ChartDataSourceSpec.StringLiteral(List.of("Jan"))));

      Name sheetScoped = workbook.createName();
      sheetScoped.setNameName("SheetScopedValues");
      sheetScoped.setSheetIndex(workbook.getSheetIndex(sheet));
      sheetScoped.setRefersToFormula("B2:B3");
      assertEquals(sheet, ChartAdapter.resolveAreaReference(sheet, "SheetScopedValues").sheet());

      Name workbookScoped = workbook.createName();
      workbookScoped.setNameName("WorkbookValues");
      workbookScoped.setRefersToFormula("Charts!$B$2:$B$3");
      assertEquals(
          "WorkbookValues",
          ChartAdapter.resolveDefinedNameReference(sheet, "WorkbookValues")
              .orElseThrow()
              .getNameName());
      assertEquals(
          List.of("10.0", "12.0"),
          ChartAdapter.resolveChartSource(sheet, "WorkbookValues").stringValues());

      try (XssfWorkbook wrappedWorkbook = XssfWorkbooks.wrap(workbook)) {
        XssfSheet wrappedSheet = wrappedWorkbook.getOrCreateSheet("Ops");
        wrappedSheet
            .drawings()
            .setSignatureLine(
                new ResolvedSignatureLine(
                    "Dup",
                    ExcelChartTestSupport.anchor(5, 1, 8, 6),
                    false,
                    Optional.ofNullable("instructions"),
                    Optional.ofNullable("Ada"),
                    Optional.ofNullable("Finance"),
                    Optional.ofNullable("ada@example.com"),
                    Optional.empty(),
                    Optional.ofNullable("invalid"),
                    Optional.of(ExcelPictureFormat.PNG),
                    Optional.of(new BinaryContent(new byte[] {1}))));
        forceSignatureLineName(wrappedWorkbook.sheet("Ops").xssfSheet(), "Dup");
        XSSFDrawing drawing = wrappedWorkbook.sheet("Ops").xssfSheet().createDrawingPatriarch();
        XSSFSimpleShape duplicateShape =
            drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 1, 1, 4, 6));
        duplicateShape.getCTShape().getNvSpPr().getCNvPr().setName("Dup");
        duplicateShape.setText("text");
        DrawingAdapter controller = new DrawingAdapter();
        assertTrue(
            new SignatureAdapter()
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

  private static CustomXmlMappingReport customXmlMappingSnapshot(
      long mapId,
      String name,
      String rootElement,
      String schemaId,
      CustomXmlMappingReport.Settings settings,
      CustomXmlMappingReport.Schema schema,
      Optional<CustomXmlDataBindingReport> dataBinding,
      List<CustomXmlLinkedCellReport> linkedCells,
      List<CustomXmlLinkedTableReport> linkedTables) {
    return new CustomXmlMappingReport(
        mapId,
        name,
        rootElement,
        schemaId,
        settings,
        schema,
        dataBinding == null ? null : dataBinding.orElse(null),
        linkedCells,
        linkedTables);
  }

  private static CustomXmlMappingReport.Settings settings(
      boolean showImportExportValidationErrors,
      boolean autoFit,
      boolean append,
      boolean preserveSortAfLayout,
      boolean preserveFormat) {
    return new CustomXmlMappingReport.Settings(
        showImportExportValidationErrors, autoFit, append, preserveSortAfLayout, preserveFormat);
  }

  private static CustomXmlMappingReport.Schema schema(
      Optional<String> namespace,
      Optional<String> language,
      Optional<String> reference,
      Optional<String> xml) {
    return new CustomXmlMappingReport.Schema(
        nullable(namespace), nullable(language), nullable(reference), nullable(xml));
  }

  private static <T> T nullable(Optional<T> value) {
    return value == null ? null : value.orElse(null);
  }
}
