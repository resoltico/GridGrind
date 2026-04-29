package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.query.InspectionQuery;
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
import org.junit.jupiter.api.Test;

/**
 * Covers remaining XLSX protocol DTO validation branches for charts, signatures, and custom XML.
 */
class SpreadsheetSurfaceEdgeCoverageTest {
  @Test
  void chartInputFamiliesNormalizeExtendedPlotBranches() {
    ChartInput.Axis defaultVisibleAxis =
        ChartInput.Axis.create(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            null);
    ChartInput.Series series =
        ChartInput.Series.create(
            null,
            new ChartInput.DataSource.StringLiteral(List.of("Jan", "Feb")),
            new ChartInput.DataSource.NumericLiteral(List.of(10.0d, 18.0d)),
            true,
            ExcelChartMarkerStyle.DIAMOND,
            (short) 8,
            12L);

    ChartInput.Area area = ChartInput.Area.create(null, null, null, List.of(series));
    ChartInput.Area3D area3D = ChartInput.Area3D.create(null, null, 42, null, List.of(series));
    ChartInput.Bar3D bar3D =
        ChartInput.Bar3D.create(
            null,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.PERCENT_STACKED,
            24,
            88,
            ExcelChartBarShape.CONE,
            null,
            List.of(series));
    ChartInput.Doughnut doughnut = ChartInput.Doughnut.create(null, 45, 40, List.of(series));
    ChartInput.Line3D line3D = ChartInput.Line3D.create(null, null, 16, null, List.of(series));
    ChartInput.Pie3D pie3D = ChartInput.Pie3D.create(null, List.of(series));
    ChartInput.Radar radar = ChartInput.Radar.create(null, null, null, List.of(series));
    ChartInput.Scatter scatter = ChartInput.Scatter.create(null, null, null, List.of(series));
    ChartInput.Surface surface = ChartInput.Surface.create(null, null, null, List.of(series));
    ChartInput.Surface3D surface3D = ChartInput.Surface3D.create(null, null, null, List.of(series));

    assertTrue(defaultVisibleAxis.visible());
    assertTrue(series.title() instanceof ChartInput.Title.None);
    assertEquals(
        List.of("Jan", "Feb"),
        ((ChartInput.DataSource.StringLiteral) series.categories()).values());
    assertEquals(
        List.of(10.0d, 18.0d), ((ChartInput.DataSource.NumericLiteral) series.values()).values());
    assertFalse(area.varyColors());
    assertEquals(ExcelChartGrouping.STANDARD, area.grouping());
    assertEquals(2, area.axes().size());
    assertEquals(42, area3D.gapDepth());
    assertFalse(area3D.varyColors());
    assertEquals(ExcelChartBarDirection.BAR, bar3D.barDirection());
    assertEquals(ExcelChartBarGrouping.PERCENT_STACKED, bar3D.grouping());
    assertEquals(ExcelChartBarShape.CONE, bar3D.shape());
    assertFalse(doughnut.varyColors());
    assertEquals(45, doughnut.firstSliceAngle());
    assertEquals(40, doughnut.holeSize());
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
            ChartInput.Series.create(
                null,
                new ChartInput.DataSource.Reference("Categories"),
                new ChartInput.DataSource.Reference("Values"),
                null,
                null,
                (short) 1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ChartInput.Series.create(
                null,
                new ChartInput.DataSource.Reference("Categories"),
                new ChartInput.DataSource.Reference("Values"),
                null,
                null,
                (short) 73,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ChartInput.Series.create(
                null,
                new ChartInput.DataSource.Reference("Categories"),
                new ChartInput.DataSource.Reference("Values"),
                null,
                null,
                null,
                -1L));
    assertThrows(
        NullPointerException.class,
        () -> new ChartInput.DataSource.StringLiteral(List.of("Jan", null)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartInput.DataSource.NumericLiteral(List.of(1.0d, null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> ChartInput.Doughnut.create(false, 0, 9, List.of(series)));
    assertThrows(
        IllegalArgumentException.class,
        () -> ChartInput.Doughnut.create(false, 0, 91, List.of(series)));
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
                "ChartValues", "#,##0.00", List.of("10", "18")),
            true,
            ExcelChartMarkerStyle.SQUARE,
            (short) 6,
            4L);
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
            false, ExcelChartGrouping.PERCENT_STACKED, 24, axes, List.of(series));
    ChartReport.Bar3D bar3D =
        new ChartReport.Bar3D(
            true,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.STACKED,
            32,
            88,
            ExcelChartBarShape.PYRAMID,
            axes,
            List.of(series));
    ChartReport.Doughnut doughnut = new ChartReport.Doughnut(true, 30, 55, List.of(series));
    ChartReport.Line3D line3D =
        new ChartReport.Line3D(false, ExcelChartGrouping.STANDARD, 18, axes, List.of(series));
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
        new DrawingObjectReport.SignatureLine(
            "Signature",
            anchor,
            "{ABC}",
            true,
            "Review before signing.",
            "Ada Lovelace",
            "Finance",
            "ada@example.com",
            ExcelPictureFormat.PNG,
            "image/png",
            42L,
            "sig123",
            400,
            150);
    CustomXmlMappingReport mapping =
        new CustomXmlMappingReport(
            1L,
            "CORSO_mapping",
            "CORSO",
            "Schema1",
            false,
            true,
            false,
            true,
            true,
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
    assertEquals(24, area3D.gapDepth());
    assertEquals(ExcelChartBarShape.PYRAMID, bar3D.shape());
    assertEquals(55, doughnut.holeSize());
    assertEquals(18, line3D.gapDepth());
    assertTrue(pie3D.varyColors());
    assertEquals(ExcelChartRadarStyle.FILLED, radar.style());
    assertEquals(ExcelChartScatterStyle.SMOOTH_MARKER, scatter.style());
    assertTrue(surface.wireframe());
    assertFalse(surface3D.wireframe());
    assertEquals("{ABC}", signatureLine.setupId());
    assertEquals("image/png", signatureLine.previewContentType());
    assertEquals("urn:gridgrind:test", mapping.schemaNamespace());
    assertEquals("binding.xml", mapping.dataBinding().fileBindingName());
    assertEquals("CourseTable", mapping.linkedTables().getFirst().tableDisplayName());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                " ",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                " ",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                " ",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                " ",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                " ",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                "image/png",
                42L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                " ",
                42L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                null,
                null,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                -1L,
                "sig123",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                null,
                null,
                " ",
                10,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "sig123",
                -1,
                10));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                anchor,
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "sig123",
                10,
                -1));
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
        () ->
            new CustomXmlMappingReport(
                1L,
                "CORSO_mapping",
                "CORSO",
                "Schema1",
                false,
                true,
                false,
                true,
                true,
                " ",
                null,
                null,
                null,
                null,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlLinkedTableReport("Foglio1", "Table1", " ", "A1:B4", "/rows/row"));
    assertThrows(
        NullPointerException.class,
        () ->
            new CustomXmlMappingReport(
                1L,
                "CORSO_mapping",
                "CORSO",
                "Schema1",
                false,
                true,
                false,
                true,
                true,
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
            new CustomXmlMappingReport(
                1L,
                "CORSO_mapping",
                "CORSO",
                "Schema1",
                false,
                true,
                false,
                true,
                true,
                null,
                null,
                null,
                null,
                null,
                List.of(),
                List.of((CustomXmlLinkedTableReport) null)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartReport.Area3D(false, null, 12, axes, List.of(series)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartReport.Bar3D(
                false, ExcelChartBarDirection.COLUMN, null, 12, 44, null, axes, List.of(series)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartReport.Doughnut(false, 0, 9, List.of(series)));
    assertThrows(
        NullPointerException.class,
        () -> new ChartReport.Line3D(false, null, 12, axes, List.of(series)));
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
    InspectionQuery.ExportCustomXmlMapping export =
        new InspectionQuery.ExportCustomXmlMapping(
            new CustomXmlMappingLocator(1L, null), false, "UTF-8");

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
            new InspectionQuery.ExportCustomXmlMapping(
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
            ChartInput.Bar.create(
                null,
                null,
                null,
                null,
                101,
                null,
                List.of(
                    ChartInput.Series.create(
                        null,
                        new ChartInput.DataSource.StringLiteral(List.of("Jan")),
                        new ChartInput.DataSource.NumericLiteral(List.of(10.0d)),
                        null,
                        null,
                        null,
                        null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ChartInput.Bar.create(
                null,
                null,
                null,
                null,
                -101,
                null,
                List.of(
                    ChartInput.Series.create(
                        null,
                        new ChartInput.DataSource.StringLiteral(List.of("Jan")),
                        new ChartInput.DataSource.NumericLiteral(List.of(10.0d)),
                        null,
                        null,
                        null,
                        null))));
    assertEquals(
        ExcelChartBarDirection.COLUMN,
        ChartInput.Bar3D.create(
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                List.of(
                    ChartInput.Series.create(
                        null,
                        new ChartInput.DataSource.StringLiteral(List.of("Jan")),
                        new ChartInput.DataSource.NumericLiteral(List.of(10.0d)),
                        null,
                        null,
                        null,
                        null)))
            .barDirection());
    assertNull(
        ChartInput.Doughnut.create(
                null,
                0,
                null,
                List.of(
                    ChartInput.Series.create(
                        null,
                        new ChartInput.DataSource.StringLiteral(List.of("Jan")),
                        new ChartInput.DataSource.NumericLiteral(List.of(10.0d)),
                        null,
                        null,
                        null,
                        null)))
            .holeSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ChartInput.Doughnut.create(
                null,
                0,
                91,
                List.of(
                    ChartInput.Series.create(
                        null,
                        new ChartInput.DataSource.StringLiteral(List.of("Jan")),
                        new ChartInput.DataSource.NumericLiteral(List.of(10.0d)),
                        null,
                        null,
                        null,
                        null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(null, List.of("10")),
                null,
                null,
                (short) 1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(null, List.of("10")),
                null,
                null,
                (short) 73,
                null));
    assertNull(
        new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(null, List.of("10")),
                null,
                null,
                null,
                null)
            .markerSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Series"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(null, List.of("10")),
                null,
                null,
                null,
                -1L));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Doughnut(
                false,
                0,
                91,
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Series"),
                        new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                        new ChartReport.DataSource.NumericLiteral(null, List.of("10")),
                        null,
                        null,
                        null,
                        null))));
    assertNull(
        new ChartReport.Doughnut(
                false,
                0,
                null,
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Series"),
                        new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                        new ChartReport.DataSource.NumericLiteral(null, List.of("10")),
                        null,
                        null,
                        null,
                        null)))
            .holeSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                null,
                null,
                "sig123",
                10,
                10));
    assertEquals(
        "preview",
        new DrawingObjectReport.SignatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                "preview",
                null,
                null)
            .previewSha256());
    assertNull(
        new DrawingObjectReport.SignatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                null,
                null,
                null)
            .previewSha256());
    assertNull(
        new DrawingObjectReport.SignatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                null,
                null,
                null,
                null,
                null,
                null)
            .previewSha256());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.SignatureLine(
                "Signature",
                new DrawingAnchorReport.TwoCell(
                    new DrawingMarkerReport(1, 2, 0, 0),
                    new DrawingMarkerReport(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                "{ABC}",
                true,
                "Review",
                "Ada",
                "Finance",
                "ada@example.com",
                ExcelPictureFormat.PNG,
                "image/png",
                42L,
                " ",
                10,
                10));
    assertEquals(
        null,
        new CustomXmlMappingReport(
                1L,
                "CORSO_mapping",
                "CORSO",
                "Schema1",
                false,
                true,
                false,
                true,
                true,
                null,
                null,
                null,
                null,
                null,
                List.of(new CustomXmlLinkedCellReport("Foglio1", "A1", "/CORSO/NOME", "string")),
                List.of())
            .schemaXml());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CustomXmlMappingReport(
                1L,
                "CORSO_mapping",
                "CORSO",
                "Schema1",
                false,
                true,
                false,
                true,
                true,
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

  private static List<ExcelChartAxisKind> kinds(List<? extends ChartInput.Axis> axes) {
    return axes.stream().map(ChartInput.Axis::kind).toList();
  }
}
