package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/**
 * Covers remaining XLSX protocol DTO validation branches for charts, signatures, and custom XML.
 */
class SpreadsheetSurfaceEdgeCoverageTest {
  @Test
  void chartInputFamiliesNormalizeExtendedPlotBranches() {
    ChartAxisInput defaultVisibleAxis =
        ChartAxisInput.visible(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO);
    ChartSeriesInput series =
        ChartSeriesInput.untitled(
            new ChartDataSourceInput.StringLiteral(List.of("Jan", "Feb")),
            new ChartDataSourceInput.NumericLiteral(List.of(10.0d, 18.0d)),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.DIAMOND),
            Optional.of((short) 8),
            Optional.of(12L));

    ChartPlotInput.Area area =
        new ChartPlotInput.Area(false, ExcelChartGrouping.STANDARD, List.of(series));
    ChartPlotInput.Area3D area3D =
        new ChartPlotInput.Area3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(42), List.of(series));
    ChartPlotInput.Bar3D bar3D =
        new ChartPlotInput.Bar3D(
            false,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.PERCENT_STACKED,
            Optional.of(24),
            Optional.of(88),
            Optional.of(ExcelChartBarShape.CONE),
            List.of(series));
    ChartPlotInput.Doughnut doughnut =
        new ChartPlotInput.Doughnut(false, Optional.of(45), Optional.of(40), List.of(series));
    ChartPlotInput.Line3D line3D =
        new ChartPlotInput.Line3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(16), List.of(series));
    ChartPlotInput.Pie3D pie3D = new ChartPlotInput.Pie3D(false, List.of(series));
    ChartPlotInput.Radar radar =
        new ChartPlotInput.Radar(false, ExcelChartRadarStyle.STANDARD, List.of(series));
    ChartPlotInput.Scatter scatter =
        new ChartPlotInput.Scatter(false, ExcelChartScatterStyle.LINE_MARKER, List.of(series));
    ChartPlotInput.Surface surface = new ChartPlotInput.Surface(false, false, List.of(series));
    ChartPlotInput.Surface3D surface3D =
        new ChartPlotInput.Surface3D(false, false, List.of(series));

    assertTrue(defaultVisibleAxis.visible());
    assertTrue(series.title() instanceof ChartTitleInput.None);
    assertEquals(
        List.of("Jan", "Feb"), ((ChartDataSourceInput.StringLiteral) series.categories()).values());
    assertEquals(
        List.of(10.0d, 18.0d), ((ChartDataSourceInput.NumericLiteral) series.values()).values());
    assertFalse(area.varyColors());
    assertEquals(ExcelChartGrouping.STANDARD, area.grouping());
    assertEquals(2, area.axes().size());
    assertEquals(Optional.of(42), area3D.gapDepth());
    assertFalse(area3D.varyColors());
    assertEquals(ExcelChartBarDirection.BAR, bar3D.barDirection());
    assertEquals(ExcelChartBarGrouping.PERCENT_STACKED, bar3D.grouping());
    assertEquals(Optional.of(ExcelChartBarShape.CONE), bar3D.shape());
    assertFalse(doughnut.varyColors());
    assertEquals(Optional.of(45), doughnut.firstSliceAngle());
    assertEquals(Optional.of(40), doughnut.holeSize());
    assertEquals(ExcelChartGrouping.STANDARD, line3D.grouping());
    assertFalse(pie3D.varyColors());
    assertEquals(ExcelChartRadarStyle.STANDARD, radar.style());
    assertEquals(ExcelChartScatterStyle.LINE_MARKER, scatter.style());
    assertEquals(
        List.of(ExcelChartAxisKind.VALUE, ExcelChartAxisKind.VALUE), kinds(scatter.axes()));
    assertFalse(surface.varyColors());
    assertFalse(surface.wireframe());
    assertEquals(
        List.of(ExcelChartAxisKind.CATEGORY, ExcelChartAxisKind.VALUE, ExcelChartAxisKind.SERIES),
        kinds(surface.axes()));
    assertFalse(surface3D.varyColors());
    assertFalse(surface3D.wireframe());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            ChartSeriesInput.untitled(
                new ChartDataSourceInput.Reference("Categories"),
                new ChartDataSourceInput.Reference("Values"),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ChartSeriesInput.untitled(
                new ChartDataSourceInput.Reference("Categories"),
                new ChartDataSourceInput.Reference("Values"),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 73),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ChartSeriesInput.untitled(
                new ChartDataSourceInput.Reference("Categories"),
                new ChartDataSourceInput.Reference("Values"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1L)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartDataSourceInput.StringLiteral(List.of("Jan", null)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartDataSourceInput.NumericLiteral(List.of(1.0d, null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartPlotInput.Doughnut(false, Optional.of(0), Optional.of(9), List.of(series)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartPlotInput.Doughnut(false, Optional.of(0), Optional.of(91), List.of(series)));
  }

  @Test
  void chartReportCustomXmlAndSignatureReportsValidateExtendedBranches() {
    DrawingAnchorReport.TwoCell anchor =
        new DrawingAnchorReport.TwoCell(
            new DrawingMarkerReport(1, 2, 0, 0),
            new DrawingMarkerReport(6, 12, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ChartReport.Series series =
        new ChartReport.Series(
            new ChartReport.Title.Text("Series"),
            new ChartReport.DataSource.StringReference("ChartCategories", List.of("Jan", "Feb")),
            new ChartReport.DataSource.NumericReference(
                "ChartValues", Optional.of("#,##0.00"), List.of("10", "18")),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.SQUARE),
            Optional.of((short) 6),
            Optional.of(4L));
    List<ChartReport.Axis> axes =
        List.of(
            new ChartReport.Axis(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ChartReport.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true));

    ChartReport.Area area =
        new ChartReport.Area(false, ExcelChartGrouping.STANDARD, axes, List.of(series));
    ChartReport.Area3D area3D =
        new ChartReport.Area3D(
            false, ExcelChartGrouping.PERCENT_STACKED, Optional.of(24), axes, List.of(series));
    ChartReport.Bar3D bar3D =
        new ChartReport.Bar3D(
            true,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.STACKED,
            Optional.of(32),
            Optional.of(88),
            Optional.of(ExcelChartBarShape.PYRAMID),
            axes,
            List.of(series));
    ChartReport.Doughnut doughnut =
        new ChartReport.Doughnut(true, Optional.of(30), Optional.of(55), List.of(series));
    ChartReport.Line3D line3D =
        new ChartReport.Line3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(18), axes, List.of(series));
    ChartReport.Pie3D pie3D = new ChartReport.Pie3D(true, List.of(series));
    ChartReport.Radar radar =
        new ChartReport.Radar(false, ExcelChartRadarStyle.FILLED, axes, List.of(series));
    ChartReport.Scatter scatter =
        new ChartReport.Scatter(false, ExcelChartScatterStyle.SMOOTH_MARKER, axes, List.of(series));
    ChartReport.Surface surface =
        new ChartReport.Surface(false, true, surfaceAxes(), List.of(series));
    ChartReport.Surface3D surface3D =
        new ChartReport.Surface3D(true, false, surfaceAxes(), List.of(series));
    DrawingObjectReport.SignatureLine signatureLine =
        signatureLine(
            "Signature",
            anchor,
            signatureSetup(
                "{ABC}",
                true,
                "Review before signing.",
                "Ada Lovelace",
                "Finance",
                "ada@example.com"),
            signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", 400, 150));
    CustomXmlMappingReport mapping =
        mappingReport(
            "urn:gridgrind:test",
            "XSD",
            "schema.xsd",
            "<xsd:schema/>",
            new CustomXmlDataBindingReport("binding", true, 5L, "binding.xml", 2L),
            List.of(new CustomXmlLinkedCellReport("Foglio1", "A1", "/CORSO/NOME", "string")),
            List.of(
                new CustomXmlLinkedTableReport(
                    "Foglio1", "Table1", "CourseTable", "A1:B4", "/CORSO/RIGHE/RIGA")));

    assertFalse(area.varyColors());
    assertEquals(Optional.of(24), area3D.gapDepth());
    assertEquals(Optional.of(ExcelChartBarShape.PYRAMID), bar3D.shape());
    assertEquals(Optional.of(55), doughnut.holeSize());
    assertEquals(Optional.of(18), line3D.gapDepth());
    assertTrue(pie3D.varyColors());
    assertEquals(ExcelChartRadarStyle.FILLED, radar.style());
    assertEquals(ExcelChartScatterStyle.SMOOTH_MARKER, scatter.style());
    assertTrue(surface.wireframe());
    assertFalse(surface3D.wireframe());
    assertEquals("{ABC}", signatureLine.setup().orElseThrow().setupId().orElseThrow());
    assertEquals("image/png", signatureLine.preview().orElseThrow().contentType());
    assertEquals("urn:gridgrind:test", mapping.schema().namespace());
    assertEquals("binding.xml", mapping.dataBinding().fileBindingName());
    assertEquals("CourseTable", mapping.linkedTables().getFirst().tableDisplayName());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup(" ", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, " ", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", " ", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", " ", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", " "),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(null, "image/png", 42L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, " ", 42L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(null, null, null, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", -1L, "sig123", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(null, null, null, " ", 10, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", -1, 10)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                anchor,
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "sig123", 10, -1)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlDataBindingReport("binding", true, -1L, "binding.xml", 2L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlDataBindingReport("binding", true, 1L, " ", 2L));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlDataBindingReport("binding", true, 1L, "binding.xml", -1L));
    assertThrows(
        IllegalArgumentException.class,
        () -> mappingReport(" ", null, null, null, null, List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlLinkedTableReport("Foglio1", "Table1", " ", "A1:B4", "/rows/row"));
    assertThrows(
        NullPointerException.class,
        () ->
            mappingReport(
                null,
                null,
                null,
                null,
                null,
                List.of((CustomXmlLinkedCellReport) null),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            mappingReport(
                null,
                null,
                null,
                null,
                null,
                List.of(),
                List.of((CustomXmlLinkedTableReport) null)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartReport.Area3D(false, null, Optional.of(12), axes, List.of(series)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartReport.Bar3D(
                false,
                ExcelChartBarDirection.COLUMN,
                null,
                Optional.of(12),
                Optional.of(44),
                Optional.empty(),
                axes,
                List.of(series)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartReport.Doughnut(false, Optional.of(0), Optional.of(9), List.of(series)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartReport.Line3D(false, null, Optional.of(12), axes, List.of(series)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartReport.Radar(false, null, axes, List.of(series)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartReport.Scatter(false, null, axes, List.of(series)));
  }

  @Test
  void signatureLineInputAndCustomXmlExportQueryValidateDefaultingBranches() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 0, 0),
            new DrawingMarkerInput(6, 12, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    SignatureLineInput captionOnly =
        new SignatureLineInput(
            "ApprovalSignature",
            anchor,
            true,
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of("Please sign\nbefore release"),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    SignatureLineInput signerOnly =
        new SignatureLineInput(
            "SignerOnly",
            anchor,
            false,
            java.util.Optional.of("Review before signing."),
            java.util.Optional.of("Ada Lovelace"),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of("invalid"),
            java.util.Optional.empty());
    WorkbookIntrospectionQuery.ExportCustomXmlMapping export =
        new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
            new CustomXmlMappingLocator(1L, null), false);

    assertTrue(captionOnly.allowComments());
    assertEquals("Please sign\nbefore release", captionOnly.caption().orElseThrow());
    assertFalse(signerOnly.allowComments());
    assertEquals("Ada Lovelace", signerOnly.suggestedSigner().orElseThrow());
    assertFalse(export.validateSchema());
    assertEquals("UTF-8", export.encoding());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SignatureLineInput(
                "TooManyLines",
                anchor,
                true,
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of("one\ntwo\nthree\nfour"),
                java.util.Optional.empty(),
                java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SignatureLineInput(
                "MissingSigner",
                anchor,
                true,
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SignatureLineInput(
                "BlankCaption",
                anchor,
                true,
                java.util.Optional.empty(),
                java.util.Optional.of("Ada"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of(" "),
                java.util.Optional.empty(),
                java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SignatureLineInput(
                "BlankInvalidStamp",
                anchor,
                true,
                java.util.Optional.empty(),
                java.util.Optional.of("Ada"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of(" "),
                java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                new CustomXmlMappingLocator(1L, null), true, " "));
  }

  @Test
  void miscellaneousSpreadsheetDtosCoverRemainingNormalizationBranches() {
    ArrayFormulaInput arrayFormula =
        new ArrayFormulaInput(TextSourceInput.inline("{=B2:B4*C2:C4}"));
    ArrayFormulaInput equalsFormula = new ArrayFormulaInput(TextSourceInput.inline("=SUM(A1:A2)"));
    ArrayFormulaInput malformedBraceFormula =
        new ArrayFormulaInput(TextSourceInput.inline("{=BROKEN"));
    ArrayFormulaInput standardInputFormula = new ArrayFormulaInput(TextSourceInput.standardInput());
    CustomXmlMappingLocator namedLocator = new CustomXmlMappingLocator(null, "CORSO_mapping");

    assertEquals(
        "B2:B4*C2:C4",
        assertInstanceOf(TextSourceInput.Inline.class, arrayFormula.source()).text());
    assertEquals(
        "SUM(A1:A2)",
        assertInstanceOf(TextSourceInput.Inline.class, equalsFormula.source()).text());
    assertEquals(
        "{=BROKEN",
        assertInstanceOf(TextSourceInput.Inline.class, malformedBraceFormula.source()).text());
    assertTrue(standardInputFormula.source() instanceof TextSourceInput.StandardInput);
    assertEquals("CORSO_mapping", namedLocator.name());
    assertThrows(IllegalArgumentException.class, () -> new CustomXmlMappingLocator(null, " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlLinkedCellReport("Foglio1", " ", "/CORSO/NOME", "string"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SignatureLineInput(
                " ",
                new DrawingAnchorInput.TwoCell(
                    new DrawingMarkerInput(1, 2, 0, 0),
                    new DrawingMarkerInput(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                true,
                java.util.Optional.empty(),
                java.util.Optional.of("Ada Lovelace"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartPlotInput.Bar(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.of(101),
                List.of(
                    ChartSeriesInput.untitled(
                        new ChartDataSourceInput.StringLiteral(List.of("Jan")),
                        new ChartDataSourceInput.NumericLiteral(List.of(10.0d)),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartPlotInput.Bar(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.of(-101),
                List.of(
                    ChartSeriesInput.untitled(
                        new ChartDataSourceInput.StringLiteral(List.of("Jan")),
                        new ChartDataSourceInput.NumericLiteral(List.of(10.0d)),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))));
    assertEquals(
        ExcelChartBarDirection.COLUMN,
        new ChartPlotInput.Bar3D(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                List.of(
                    ChartSeriesInput.untitled(
                        new ChartDataSourceInput.StringLiteral(List.of("Jan")),
                        new ChartDataSourceInput.NumericLiteral(List.of(10.0d)),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))
            .barDirection());
    assertEquals(
        Optional.empty(),
        new ChartPlotInput.Doughnut(
                false,
                Optional.of(0),
                Optional.empty(),
                List.of(
                    ChartSeriesInput.untitled(
                        new ChartDataSourceInput.StringLiteral(List.of("Jan")),
                        new ChartDataSourceInput.NumericLiteral(List.of(10.0d)),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))
            .holeSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartPlotInput.Doughnut(
                false,
                Optional.of(0),
                Optional.of(91),
                List.of(
                    ChartSeriesInput.untitled(
                        new ChartDataSourceInput.StringLiteral(List.of("Jan")),
                        new ChartDataSourceInput.NumericLiteral(List.of(10.0d)),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10")),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10")),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 73),
                Optional.empty()));
    assertEquals(
        Optional.empty(),
        new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10")),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())
            .markerSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10")),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1L)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Doughnut(
                false,
                Optional.of(0),
                Optional.of(91),
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Series"),
                        new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                        new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10")),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))));
    assertEquals(
        Optional.empty(),
        new ChartReport.Doughnut(
                false,
                Optional.of(0),
                Optional.empty(),
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Series"),
                        new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                        new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10")),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))
            .holeSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(null, null, null, "sig123", 10, 10)));
    assertEquals(
        "preview",
        signatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, "preview", null, null))
            .preview()
            .orElseThrow()
            .sha256()
            .orElseThrow());
    assertTrue(
        signatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, null, null, null))
            .preview()
            .orElseThrow()
            .sha256()
            .isEmpty());
    assertTrue(
        signatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(null, null, null, null, null, null))
            .preview()
            .isEmpty());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                signatureSetup("{ABC}", true, "Review", "Ada", "Finance", "ada@example.com"),
                signaturePreview(ExcelPictureFormat.PNG, "image/png", 42L, " ", 10, 10)));
    assertEquals(
        null,
        mappingReport(
                null,
                null,
                null,
                null,
                null,
                List.of(new CustomXmlLinkedCellReport("Foglio1", "A1", "/CORSO/NOME", "string")),
                List.of())
            .schema()
            .xml());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            mappingReport(
                null,
                null,
                null,
                " ",
                null,
                List.of(new CustomXmlLinkedCellReport("Foglio1", "A1", "/CORSO/NOME", "string")),
                List.of()));
    assertEquals(
        "Grace Hopper",
        new SignatureLineInput(
                "SuggestedSigner2Only",
                new DrawingAnchorInput.TwoCell(
                    new DrawingMarkerInput(1, 2, 0, 0),
                    new DrawingMarkerInput(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                true,
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of("Grace Hopper"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty())
            .suggestedSigner2()
            .orElseThrow());
    assertEquals(
        "ada@example.com",
        new SignatureLineInput(
                "SuggestedEmailOnly",
                new DrawingAnchorInput.TwoCell(
                    new DrawingMarkerInput(1, 2, 0, 0),
                    new DrawingMarkerInput(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                true,
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of("ada@example.com"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty())
            .suggestedSignerEmail()
            .orElseThrow());
  }

  private static List<ChartReport.Axis> surfaceAxes() {
    return List.of(
        new ChartReport.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartReport.Axis(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartAxisKind> kinds(List<? extends ChartAxisInput> axes) {
    return axes.stream().map(ChartAxisInput::kind).toList();
  }

  private static DrawingObjectReport.SignatureLine signatureLine(
      String name,
      DrawingAnchorReport anchor,
      Optional<DrawingObjectReport.SignatureSetup> setup,
      Optional<DrawingObjectReport.SignaturePreview> preview) {
    return DrawingObjectReportTestSupport.signatureLine(name, anchor, setup, preview);
  }

  private static Optional<DrawingObjectReport.SignatureSetup> signatureSetup(
      String setupId,
      Boolean allowComments,
      String signingInstructions,
      String suggestedSigner,
      String suggestedSigner2,
      String suggestedSignerEmail) {
    return DrawingObjectReportTestSupport.signatureSetup(
        setupId,
        allowComments,
        signingInstructions,
        suggestedSigner,
        suggestedSigner2,
        suggestedSignerEmail);
  }

  private static Optional<DrawingObjectReport.SignaturePreview> signaturePreview(
      ExcelPictureFormat format,
      String contentType,
      Long byteSize,
      String sha256,
      Integer widthPixels,
      Integer heightPixels) {
    return DrawingObjectReportTestSupport.signaturePreview(
        format, contentType, byteSize, sha256, widthPixels, heightPixels);
  }

  private static CustomXmlMappingReport mappingReport(
      String schemaNamespace,
      String schemaLanguage,
      String schemaReference,
      String schemaXml,
      CustomXmlDataBindingReport dataBinding,
      List<CustomXmlLinkedCellReport> linkedCells,
      List<CustomXmlLinkedTableReport> linkedTables) {
    return new CustomXmlMappingReport(
        1L,
        "CORSO_mapping",
        "CORSO",
        "Schema1",
        new CustomXmlMappingReport.Settings(false, true, false, true, true),
        new CustomXmlMappingReport.Schema(
            schemaNamespace, schemaLanguage, schemaReference, schemaXml),
        dataBinding,
        linkedCells,
        linkedTables);
  }
}
