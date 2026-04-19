package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import java.util.Base64;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Broad constructor and validation coverage for advanced protocol DTO families. */
class AdvancedDtoCoverageTest {
  @Test
  void drawingReportsValidatePayloadAndAnchorSemantics() {
    DrawingMarkerReport from = new DrawingMarkerReport(0, 0, 0, 0);
    DrawingMarkerReport to = new DrawingMarkerReport(1, 1, 0, 0);
    DrawingAnchorReport.TwoCell twoCell = new DrawingAnchorReport.TwoCell(from, to, null);
    DrawingAnchorReport.OneCell oneCell = new DrawingAnchorReport.OneCell(from, 1L, 2L, null);
    DrawingAnchorReport.Absolute absolute = new DrawingAnchorReport.Absolute(1L, 2L, 3L, 4L, null);

    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, twoCell.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, oneCell.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE, absolute.behavior());
    assertEquals(
        "columnIndex must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(-1, 0, 0, 0))
            .getMessage());
    assertEquals(
        "lastColumn must be greater than or equal to firstColumn",
        assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(1, 0, 0, 0))
            .getMessage());

    DrawingObjectReport.Picture picture =
        new DrawingObjectReport.Picture(
            "Picture 1",
            twoCell,
            ExcelPictureFormat.PNG,
            "image/png",
            10L,
            "abc123",
            32,
            16,
            "Preview");
    assertEquals(32, picture.widthPixels());
    assertEquals(
        "widthPixels must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.Picture(
                        "Picture 1",
                        twoCell,
                        ExcelPictureFormat.PNG,
                        "image/png",
                        10L,
                        "abc123",
                        -1,
                        16,
                        null))
            .getMessage());

    DrawingObjectReport.Shape shape =
        new DrawingObjectReport.Shape(
            "Shape 1", oneCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, "rect", "Hello", 0);
    assertEquals("Hello", shape.text());
    assertEquals(
        "childCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.Shape(
                        "Shape 1", oneCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, null, null, -1))
            .getMessage());

    DrawingObjectReport.EmbeddedObject embedded =
        new DrawingObjectReport.EmbeddedObject(
            "Object 1",
            absolute,
            ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
            "Label",
            "object.bin",
            null,
            "application/octet-stream",
            12L,
            "def456",
            null,
            null,
            null);
    assertEquals("object.bin", embedded.fileName());
    assertEquals(
        "previewByteSize requires previewFormat",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectReport.EmbeddedObject(
                        "Object 1",
                        absolute,
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        null,
                        null,
                        null,
                        "application/octet-stream",
                        12L,
                        "def456",
                        null,
                        4L,
                        null))
            .getMessage());

    String base64 = Base64.getEncoder().encodeToString(new byte[] {1, 2, 3});
    DrawingObjectPayloadReport.Picture picturePayload =
        new DrawingObjectPayloadReport.Picture(
            "Picture 1",
            ExcelPictureFormat.PNG,
            "image/png",
            "picture.png",
            "abc123",
            base64,
            null);
    assertEquals(base64, picturePayload.base64Data());
    DrawingObjectPayloadReport.EmbeddedObject embeddedPayload =
        new DrawingObjectPayloadReport.EmbeddedObject(
            "Object 1",
            ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
            "application/octet-stream",
            null,
            "sha",
            base64,
            "Label",
            "open");
    assertEquals("open", embeddedPayload.command());
    assertEquals(
        null,
        new DrawingObjectPayloadReport.EmbeddedObject(
                "Object 2",
                ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                "application/octet-stream",
                "payload.bin",
                "sha",
                base64,
                "Label",
                null)
            .command());
    assertEquals(
        "base64Data must be valid base64",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectPayloadReport.Picture(
                        "Picture 1",
                        ExcelPictureFormat.PNG,
                        "image/png",
                        "picture.png",
                        "abc123",
                        "%%%notbase64%%%",
                        null))
            .getMessage());
    assertEquals(
        "fileName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DrawingObjectPayloadReport.EmbeddedObject(
                        "Object 1",
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        "application/octet-stream",
                        " ",
                        "sha",
                        base64,
                        null,
                        null))
            .getMessage());
  }

  @Test
  void encryptionPaneAndColorReportsValidateBoundaries() {
    assertFalse(
        new OoxmlEncryptionReport(false, null, null, null, null, null, null, null).encrypted());
    assertEquals(
        "Unencrypted package reports must not include encryption detail fields",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        false, ExcelOoxmlEncryptionMode.AGILE, null, null, null, null, null, null))
            .getMessage());
    assertEquals(
        "cipherAlgorithm must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true, ExcelOoxmlEncryptionMode.AGILE, " ", "SHA-512", "CBC", 128, 16, 1000))
            .getMessage());
    assertEquals(
        "Unencrypted package reports must not include encryption detail fields",
        assertThrows(
                IllegalArgumentException.class,
                () -> new OoxmlEncryptionReport(false, null, "AES", null, null, null, null, null))
            .getMessage());
    assertEquals(
        "Unencrypted package reports must not include encryption detail fields",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(false, null, null, "SHA-512", null, null, null, null))
            .getMessage());
    assertEquals(
        "Unencrypted package reports must not include encryption detail fields",
        assertThrows(
                IllegalArgumentException.class,
                () -> new OoxmlEncryptionReport(false, null, null, null, "CBC", null, null, null))
            .getMessage());
    assertEquals(
        "Unencrypted package reports must not include encryption detail fields",
        assertThrows(
                IllegalArgumentException.class,
                () -> new OoxmlEncryptionReport(false, null, null, null, null, 128, null, null))
            .getMessage());
    assertEquals(
        "Unencrypted package reports must not include encryption detail fields",
        assertThrows(
                IllegalArgumentException.class,
                () -> new OoxmlEncryptionReport(false, null, null, null, null, null, 16, null))
            .getMessage());
    assertEquals(
        "Unencrypted package reports must not include encryption detail fields",
        assertThrows(
                IllegalArgumentException.class,
                () -> new OoxmlEncryptionReport(false, null, null, null, null, null, null, 1000))
            .getMessage());
    assertEquals(
        "mode must not be null when encrypted",
        assertThrows(
                NullPointerException.class,
                () -> new OoxmlEncryptionReport(true, null, "AES", "SHA-512", "CBC", 128, 16, 100))
            .getMessage());
    assertEquals(
        "keyBits must be positive when encrypted",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true,
                        ExcelOoxmlEncryptionMode.AGILE,
                        "AES",
                        "SHA-512",
                        "CBC",
                        null,
                        16,
                        100))
            .getMessage());
    assertEquals(
        "blockSize must be positive when encrypted",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true,
                        ExcelOoxmlEncryptionMode.AGILE,
                        "AES",
                        "SHA-512",
                        "CBC",
                        128,
                        null,
                        100))
            .getMessage());
    assertEquals(
        "spinCount must be zero or positive when encrypted",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true,
                        ExcelOoxmlEncryptionMode.AGILE,
                        "AES",
                        "SHA-512",
                        "CBC",
                        128,
                        16,
                        null))
            .getMessage());
    assertEquals(
        "hashAlgorithm must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true, ExcelOoxmlEncryptionMode.AGILE, "AES", " ", "CBC", 128, 16, 1000))
            .getMessage());
    assertEquals(
        "chainingMode must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new OoxmlEncryptionReport(
                        true, ExcelOoxmlEncryptionMode.AGILE, "AES", "SHA-512", " ", 128, 16, 1000))
            .getMessage());

    PaneInput.Frozen frozen = new PaneInput.Frozen(1, 0, 1, 0);
    assertEquals(1, frozen.leftmostColumn());
    assertEquals(
        "splitColumn and splitRow must not both be 0",
        assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(0, 0, 0, 0))
            .getMessage());
    assertEquals(
        "leftmostColumn must be 0 when splitColumn is 0: 1",
        assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(0, 1, 1, 1))
            .getMessage());
    assertEquals(
        "topRow must be 0 when splitRow is 0: 1",
        assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(1, 0, 1, 1))
            .getMessage());
    assertEquals(
        "leftmostColumn must be 0 when xSplitPosition is 0: 1",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PaneInput.Split(0, 1, 1, 0, ExcelPaneRegion.LOWER_RIGHT))
            .getMessage());
    assertEquals(0, new PaneInput.Split(1, 0, 0, 0, ExcelPaneRegion.UPPER_RIGHT).topRow());

    CellColorReport rgb = new CellColorReport("#a1b2c3");
    assertEquals("#A1B2C3", rgb.rgb());
    assertEquals(
        "color report must expose rgb, theme, or indexed semantics",
        assertThrows(
                IllegalArgumentException.class, () -> new CellColorReport(null, null, null, null))
            .getMessage());
    assertEquals(
        "tint must be finite",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellColorReport("#ABCDEF", null, null, Double.NaN))
            .getMessage());

    CellGradientFillReport gradient =
        new CellGradientFillReport(
            "linear",
            45.0d,
            null,
            null,
            null,
            null,
            List.of(
                new CellGradientStopReport(0.0d, rgb),
                new CellGradientStopReport(1.0d, new CellColorReport(null, 1, null, null))));
    assertEquals(2, gradient.stops().size());
    assertEquals(
        "stops must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellGradientFillReport("linear", null, null, null, null, null, List.of()))
            .getMessage());
    assertEquals(
        "position must be finite and between 0.0 and 1.0",
        assertThrows(IllegalArgumentException.class, () -> new CellGradientStopReport(2.0d, rgb))
            .getMessage());
    assertEquals(
        "gradient fills must not also expose foregroundColor or backgroundColor",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellFillReport(ExcelFillPattern.SOLID, null, rgb, gradient))
            .getMessage());
    assertEquals(
        "gradient fills must not also expose foregroundColor or backgroundColor",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellFillReport(ExcelFillPattern.SOLID, rgb, null, gradient))
            .getMessage());
    assertEquals(
        "fill pattern NONE does not accept colors",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellFillReport(ExcelFillPattern.NONE, rgb, null))
            .getMessage());
    assertEquals(
        "fill backgroundColor is not supported for SOLID fills",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellFillReport(ExcelFillPattern.SOLID, rgb, rgb))
            .getMessage());
    assertEquals(
        "fill pattern NONE does not accept colors",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellFillReport(ExcelFillPattern.NONE, null, rgb))
            .getMessage());
    assertEquals(
        "type must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CellGradientFillReport(
                        " ",
                        null,
                        null,
                        null,
                        null,
                        null,
                        List.of(new CellGradientStopReport(0.0d, rgb))))
            .getMessage());
    assertEquals(
        "position must be finite and between 0.0 and 1.0",
        assertThrows(
                IllegalArgumentException.class, () -> new CellGradientStopReport(Double.NaN, rgb))
            .getMessage());
    assertEquals(
        "position must be finite and between 0.0 and 1.0",
        assertThrows(IllegalArgumentException.class, () -> new CellGradientStopReport(-0.1d, rgb))
            .getMessage());
  }

  @Test
  void autofilterChartPivotAndTableReportsCoverValidAndInvalidShapes() {
    AutofilterFilterCriterionReport.Values values =
        new AutofilterFilterCriterionReport.Values(List.of("A", "B"), true);
    assertEquals(2, values.values().size());
    assertEquals(
        "conditions must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.Custom(false, List.of()))
            .getMessage());
    assertEquals(
        "type must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.Dynamic(" ", null, null))
            .getMessage());
    assertEquals(
        "value must be finite and non-negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.Top10(true, false, -1.0d, null))
            .getMessage());
    assertEquals(
        "value must be finite and non-negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AutofilterFilterCriterionReport.Top10(
                        true, false, Double.POSITIVE_INFINITY, null))
            .getMessage());
    assertEquals(
        "value must be finite when provided",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AutofilterFilterCriterionReport.Dynamic(
                        "ABOVE_AVERAGE", Double.POSITIVE_INFINITY, null))
            .getMessage());
    assertEquals(
        "maxValue must be finite when provided",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AutofilterFilterCriterionReport.Dynamic(
                        "ABOVE_AVERAGE", null, Double.POSITIVE_INFINITY))
            .getMessage());
    assertEquals(
        "iconId must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterFilterCriterionReport.Icon("3TrafficLights1", -1))
            .getMessage());

    DrawingAnchorReport.Absolute anchor =
        new DrawingAnchorReport.Absolute(
            1L, 2L, 3L, 4L, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ChartReport.Bar bar =
        new ChartReport.Bar(
            "Revenue",
            anchor,
            new ChartReport.Title.Text("Revenue"),
            new ChartReport.Legend.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            ExcelChartBarDirection.COLUMN,
            List.of(
                new ChartReport.Axis(
                    ExcelChartAxisKind.CATEGORY,
                    ExcelChartAxisPosition.BOTTOM,
                    ExcelChartAxisCrosses.AUTO_ZERO,
                    true)),
            List.of(
                new ChartReport.Series(
                    new ChartReport.Title.Text("Series 1"),
                    new ChartReport.DataSource.StringLiteral(List.of("Jan", "Feb")),
                    new ChartReport.DataSource.NumericLiteral("#,##0", List.of("1", "2")))));
    assertEquals("Revenue", bar.name());
    assertEquals(
        "firstSliceAngle must be between 0 and 360",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ChartReport.Pie(
                        "Pie",
                        anchor,
                        new ChartReport.Title.None(),
                        new ChartReport.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.ZERO,
                        false,
                        false,
                        361,
                        List.of(
                            new ChartReport.Series(
                                new ChartReport.Title.Text("Series 1"),
                                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                                new ChartReport.DataSource.NumericLiteral(null, List.of("1"))))))
            .getMessage());
    assertEquals(
        "series must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ChartReport.Line(
                        "Line",
                        anchor,
                        new ChartReport.Title.Text("Revenue"),
                        new ChartReport.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.ZERO,
                        false,
                        false,
                        List.of(),
                        List.of()))
            .getMessage());
    assertEquals(
        180,
        new ChartReport.Pie(
                "Pie",
                anchor,
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.ZERO,
                false,
                false,
                180,
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Series 1"),
                        new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                        new ChartReport.DataSource.NumericLiteral(null, List.of("1")))))
            .firstSliceAngle());
    assertEquals(
        null,
        new ChartReport.Pie(
                "Pie",
                anchor,
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.ZERO,
                false,
                false,
                null,
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Series 1"),
                        new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                        new ChartReport.DataSource.NumericLiteral(null, List.of("1")))))
            .firstSliceAngle());

    PivotTableReport.Supported pivot =
        new PivotTableReport.Supported(
            "Sales Pivot 2026",
            "Report",
            new PivotTableReport.Anchor("C5", "C5:G9"),
            new PivotTableReport.Source.Range("Data", "A1:D5"),
            List.of(new PivotTableReport.Field(0, "Region")),
            List.of(new PivotTableReport.Field(1, "Stage")),
            List.of(new PivotTableReport.Field(2, "Owner")),
            List.of(
                new PivotTableReport.DataField(
                    3,
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    "#,##0.00")),
            true);
    assertEquals("Report", pivot.sheetName());
    assertEquals(
        "valueFormat must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PivotTableReport.DataField(
                        0, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", " "))
            .getMessage());

    TableColumnReport tableColumn = new TableColumnReport(1L, "Amount", null, null, null, null);
    assertEquals("", tableColumn.uniqueName());
    assertEquals(
        "id must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TableColumnReport(-1L, "Amount", null, null, null, null))
            .getMessage());

    PrintSetupReport defaults = PrintSetupReport.defaults();
    assertEquals(1, defaults.paperSize());
    assertEquals(
        "rowBreaks must not contain negative indexes",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new PrintSetupReport(
                        PrintMarginsReport.defaults(),
                        false,
                        false,
                        false,
                        1,
                        false,
                        false,
                        1,
                        false,
                        1,
                        List.of(-1),
                        List.of()))
            .getMessage());

    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
            AnalysisSeverity.WARNING,
            "Autofilter spans multiple ranges",
            "Autofilter found on disjoint ranges.",
            new GridGrindResponse.AnalysisLocationReport.Sheet("Budget"),
            List.of("A1:B4"));
    AutofilterHealthReport report = new AutofilterHealthReport(1, summary, List.of(finding));
    assertEquals(1, report.checkedAutofilterCount());
    assertEquals(
        "checkedAutofilterCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AutofilterHealthReport(-1, summary, List.of(finding)))
            .getMessage());
  }
}
