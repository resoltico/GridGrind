package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

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
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for the richer factual readback protocol types added for advanced XSSF parity. */
class AdvancedReadProtocolTypesTest {
  @Test
  void pivotReportsHealthReportsAndReadOperationsValidateRichShapes() {
    PivotTableReport.Source.Range rangeSource = new PivotTableReport.Source.Range("Data", "A1:D5");
    PivotTableReport.Source.Table tableSource =
        new PivotTableReport.Source.Table("SalesTable2026", "Data", "A1:D5");
    PivotTableReport.Supported supported =
        new PivotTableReport.Supported(
            "Sales Pivot 2026",
            "Report",
            new PivotTableReport.Anchor("C5", "C5:G9"),
            new PivotTableReport.Source.NamedRange("PivotSource", "Data", "A1:D5"),
            List.of(new PivotTableReport.Field(0, "Region")),
            List.of(new PivotTableReport.Field(1, "Stage")),
            List.of(new PivotTableReport.Field(2, "Owner")),
            List.of(
                new PivotTableReport.DataField(
                    3,
                    "Amount",
                    dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    "#,##0.00")),
            true);
    PivotTableReport.Unsupported unsupported =
        new PivotTableReport.Unsupported(
            "Broken Pivot",
            "Report",
            new PivotTableReport.Anchor("A3", "A3:C8"),
            "Pivot cache source no longer resolves cleanly.");
    PivotTableHealthReport health =
        new PivotTableHealthReport(
            1,
            new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0),
            List.of(
                new GridGrindResponse.AnalysisFindingReport(
                    AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
                    AnalysisSeverity.WARNING,
                    "Pivot table name is missing",
                    "GridGrind assigned a synthetic identifier for readback.",
                    new GridGrindResponse.AnalysisLocationReport.Sheet("Report"),
                    List.of("_GG_PIVOT_Report_A3"))));
    WorkbookReadOperation.GetPivotTables getPivotTables =
        new WorkbookReadOperation.GetPivotTables("pivots", new PivotTableSelection.All());
    WorkbookReadOperation.AnalyzePivotTableHealth analyzePivotTableHealth =
        new WorkbookReadOperation.AnalyzePivotTableHealth(
            "pivot-health", new PivotTableSelection.ByNames(List.of("Sales Pivot 2026")));
    WorkbookReadResult.PivotTablesResult pivotTablesResult =
        new WorkbookReadResult.PivotTablesResult("pivots", List.of(supported, unsupported));
    WorkbookReadResult.PivotTableHealthResult pivotTableHealthResult =
        new WorkbookReadResult.PivotTableHealthResult("pivot-health", health);

    assertTrue(supported.valuesAxisOnColumns());
    assertEquals("A1:D5", rangeSource.range());
    assertEquals("SalesTable2026", tableSource.name());
    assertEquals("PivotSource", ((PivotTableReport.Source.NamedRange) supported.source()).name());
    assertEquals("Pivot cache source no longer resolves cleanly.", unsupported.detail());
    assertEquals(1, health.checkedPivotTableCount());
    assertEquals("pivots", getPivotTables.requestId());
    assertEquals(
        List.of("Sales Pivot 2026"),
        ((PivotTableSelection.ByNames) analyzePivotTableHealth.selection()).names());
    assertEquals(2, pivotTablesResult.pivotTables().size());
    assertEquals(1, pivotTableHealthResult.analysis().summary().warningCount());
    assertThrows(IllegalArgumentException.class, () -> new PivotTableReport.Field(-1, "Region"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.DataField(
                -1,
                "Amount",
                dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                "#,##0.00"));
    assertThrows(
        NullPointerException.class,
        () -> new PivotTableReport.DataField(0, "Amount", null, "Total Amount", "#,##0.00"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.DataField(
                0,
                "Amount",
                dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
                " ",
                "#,##0.00"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.DataField(
                0,
                "Amount",
                dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                " "));
    assertThrows(
        IllegalArgumentException.class, () -> new PivotTableReport.Source.Range("Data", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableReport.Unsupported(
                "Broken Pivot", "Report", new PivotTableReport.Anchor("A3", "A3:C8"), " "));
    assertThrows(IllegalArgumentException.class, () -> new PivotTableReport.Anchor(" ", "A3:C8"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PivotTableHealthReport(
                -1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadOperation.GetPivotTables("pivots", null));
  }

  @Test
  void cellColorAndGradientReportsValidateRichReadShapes() {
    assertEquals(new CellColorReport("#ABCDEF"), new CellColorReport("#abcdef"));
    assertEquals(4, new CellColorReport(null, 4, null, 0.45d).theme());
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(" ", null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellColorReport("#12345G", null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(null, -1, null, null));
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellColorReport(null, null, null, Double.NaN));

    CellGradientStopReport firstStop = new CellGradientStopReport(0.0d, rgb("#112233"));
    CellGradientStopReport secondStop =
        new CellGradientStopReport(1.0d, new CellColorReport(null, 4, null, 0.45d));
    CellGradientFillReport gradient =
        new CellGradientFillReport(
            "LINEAR", 45.0d, 0.1d, 0.2d, 0.3d, 0.4d, List.of(firstStop, secondStop));

    assertEquals(2, gradient.stops().size());
    assertEquals(0.2d, gradient.right());
    assertEquals(
        new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), null),
        new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), null));
    assertEquals(
        new CellFillReport(ExcelFillPattern.NONE, null, null),
        new CellFillReport(ExcelFillPattern.NONE, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellGradientStopReport(1.5d, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillReport(
                " ", 45.0d, null, null, null, null, List.of(firstStop, secondStop)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillReport(
                "LINEAR", Double.POSITIVE_INFINITY, null, null, null, null, List.of(firstStop)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientFillReport("LINEAR", 45.0d, null, null, null, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellGradientFillReport(
                "LINEAR", 45.0d, null, null, null, null, Arrays.asList(firstStop, null)));

    assertEquals(
        gradient, new CellFillReport(ExcelFillPattern.NONE, null, null, gradient).gradient());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillReport(ExcelFillPattern.NONE, rgb("#112233"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), rgb("#445566")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), null, gradient));
  }

  @Test
  void commentPrintAndWorkbookProtectionReportsValidateRichReadShapes() {
    CellFontReport baseFont = font(rgb("#112233"));
    CommentAnchorReport anchor = new CommentAnchorReport(1, 2, 4, 6);
    GridGrindResponse.CommentReport plainComment =
        new GridGrindResponse.CommentReport("Plain", "Ada", false);
    GridGrindResponse.CommentReport comment =
        new GridGrindResponse.CommentReport(
            "Hi there",
            "Ada",
            true,
            List.of(
                new RichTextRunReport("Hi ", baseFont),
                new RichTextRunReport("there", font(new CellColorReport(null, 5, null, 0.2d)))),
            anchor);

    assertNull(plainComment.runs());
    assertNull(plainComment.anchor());
    assertEquals(anchor, comment.anchor());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CommentReport(
                "Mismatch", "Ada", true, List.of(new RichTextRunReport("Other", baseFont)), null));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(2, 0, 1, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(0, 2, 0, 1));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(-1, 0, 1, 1));

    PrintSetupReport setup =
        new PrintSetupReport(
            new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
            true,
            false,
            9,
            false,
            true,
            2,
            true,
            3,
            List.of(10, 20),
            List.of(2, 4));

    assertEquals(List.of(10, 20), setup.rowBreaks());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsReport(-0.1d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                true,
                false,
                -1,
                false,
                true,
                2,
                true,
                3,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                List.of(-1),
                List.of(2)));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                Arrays.asList(10, null),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                true,
                false,
                9,
                false,
                true,
                -1,
                true,
                3,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                List.of(10),
                List.of(-1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                -1,
                List.of(10),
                List.of(2)));

    WorkbookProtectionReport protection =
        new WorkbookProtectionReport(true, false, true, true, false);
    assertTrue(protection.structureLocked());
  }

  @Test
  void autofilterReportsValidateAdvancedCriteriaAndSortShapes() {
    AutofilterFilterCriterionReport.Values values =
        new AutofilterFilterCriterionReport.Values(List.of("Queued", "Blocked"), true);
    AutofilterFilterCriterionReport.Custom custom =
        new AutofilterFilterCriterionReport.Custom(
            true,
            List.of(new AutofilterFilterCriterionReport.CustomConditionReport("equal", "Ada")));
    AutofilterFilterCriterionReport.Dynamic dynamic =
        new AutofilterFilterCriterionReport.Dynamic("TODAY", 1.0d, 2.0d);
    AutofilterFilterCriterionReport.Top10 top10 =
        new AutofilterFilterCriterionReport.Top10(true, false, 10.0d, 8.0d);
    AutofilterFilterCriterionReport.Color color =
        new AutofilterFilterCriterionReport.Color(false, new CellColorReport(null, 4, null, 0.45d));
    AutofilterFilterCriterionReport.Icon icon =
        new AutofilterFilterCriterionReport.Icon("3TrafficLights1", 2);
    AutofilterSortConditionReport sortCondition =
        new AutofilterSortConditionReport("A2:A5", true, null, rgb("#AABBCC"), 1);
    AutofilterSortStateReport sortState =
        new AutofilterSortStateReport("A1:F5", true, false, null, List.of(sortCondition));
    AutofilterEntryReport.SheetOwned sheetOwned =
        new AutofilterEntryReport.SheetOwned(
            "A1:F5",
            List.of(
                new AutofilterFilterColumnReport(0L, false, values),
                new AutofilterFilterColumnReport(1L, true, custom),
                new AutofilterFilterColumnReport(2L, true, dynamic),
                new AutofilterFilterColumnReport(3L, true, top10),
                new AutofilterFilterColumnReport(4L, true, color),
                new AutofilterFilterColumnReport(5L, true, icon)),
            sortState);
    AutofilterEntryReport.TableOwned tableOwned =
        new AutofilterEntryReport.TableOwned("H1:I5", "QueueTable", List.of(), sortState);
    AutofilterEntryReport.TableOwned tableOwnedWithColumn =
        new AutofilterEntryReport.TableOwned(
            "N1:O5",
            "QueueMirror",
            List.of(new AutofilterFilterColumnReport(9L, true, values)),
            sortState);
    AutofilterEntryReport.SheetOwned defaultSheetOwned =
        new AutofilterEntryReport.SheetOwned("J1:K4");
    AutofilterEntryReport.TableOwned defaultTableOwned =
        new AutofilterEntryReport.TableOwned("L1:M4", "AuditTable");
    AutofilterFilterCriterionReport.Values emptyValues =
        new AutofilterFilterCriterionReport.Values(List.of(), false);

    assertEquals(sortState, sheetOwned.sortState());
    assertEquals("QueueTable", tableOwned.tableName());
    assertEquals(1, tableOwnedWithColumn.filterColumns().size());
    assertEquals(List.of(), emptyValues.values());
    assertEquals(List.of(), defaultSheetOwned.filterColumns());
    assertNull(defaultSheetOwned.sortState());
    assertEquals(List.of(), defaultTableOwned.filterColumns());
    assertNull(defaultTableOwned.sortState());
    assertEquals("", sortCondition.sortBy());
    assertEquals("", sortState.sortMethod());

    assertThrows(
        NullPointerException.class,
        () -> new AutofilterFilterCriterionReport.Values(Arrays.asList("Queued", null), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Custom(true, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.CustomConditionReport(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic(" ", 1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic("TODAY", Double.POSITIVE_INFINITY, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic("TODAY", 1.0d, Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Top10(true, false, -1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterFilterCriterionReport.Top10(
                true, false, 10.0d, Double.NEGATIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterCriterionReport.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterColumnReport(-1L, true, values));
    assertEquals(
        "CELL_COLOR",
        new AutofilterSortConditionReport("A2:A5", false, "CELL_COLOR", null, null).sortBy());
    assertEquals(" ", new AutofilterSortConditionReport(" ", false, null, null, null).range());
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionReport("A2:A5", true, null, rgb("#AABBCC"), -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.CustomConditionReport("equal", " "));
    assertEquals(
        " ", new AutofilterSortStateReport(" ", true, false, null, List.of(sortCondition)).range());
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterSortStateReport(
                "A1:F5", true, false, null, Arrays.asList(sortCondition, null)));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterEntryReport.SheetOwned(
                "A1:F5", Arrays.asList(sheetOwned.filterColumns().getFirst(), null), sortState));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterEntryReport.TableOwned(
                "A1:F5",
                "QueueTable",
                Arrays.asList(sheetOwned.filterColumns().getFirst(), null),
                sortState));
  }

  @Test
  void drawingReportsValidateReadShapes() {
    DrawingMarkerReport from = new DrawingMarkerReport(1, 2, 3, 4);
    DrawingMarkerReport to = new DrawingMarkerReport(4, 6, 0, 0);
    DrawingAnchorReport.TwoCell twoCell =
        new DrawingAnchorReport.TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    DrawingAnchorReport.TwoCell defaultTwoCell = new DrawingAnchorReport.TwoCell(from, to, null);
    DrawingAnchorReport.OneCell oneCell = new DrawingAnchorReport.OneCell(from, 10L, 20L, null);
    DrawingAnchorReport.Absolute absolute =
        new DrawingAnchorReport.Absolute(1L, 2L, 10L, 20L, null);
    DrawingObjectReport.Picture picture =
        new DrawingObjectReport.Picture(
            "OpsPicture",
            twoCell,
            ExcelPictureFormat.PNG,
            "image/png",
            68L,
            "abc123",
            null,
            null,
            "Queue preview");
    DrawingObjectReport.Shape shape =
        new DrawingObjectReport.Shape(
            "OpsShape", oneCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, "rect", "Queue", 0);
    DrawingObjectReport.EmbeddedObject embeddedObject =
        new DrawingObjectReport.EmbeddedObject(
            "OpsEmbed",
            absolute,
            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
            "Payload",
            "payload.txt",
            "payload.txt",
            "application/octet-stream",
            7L,
            "def456",
            null,
            null,
            null);
    DrawingObjectPayloadReport.Picture picturePayload =
        new DrawingObjectPayloadReport.Picture(
            "OpsPicture",
            ExcelPictureFormat.PNG,
            "image/png",
            "OpsPicture.png",
            "abc123",
            "cGljdHVyZQ==",
            "Queue preview");
    DrawingObjectPayloadReport.EmbeddedObject embeddedPayload =
        new DrawingObjectPayloadReport.EmbeddedObject(
            "OpsEmbed",
            ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
            "application/octet-stream",
            "payload.txt",
            "def456",
            "cGF5bG9hZA==",
            "Payload",
            "payload.txt");

    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, defaultTwoCell.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, oneCell.behavior());
    assertEquals(ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE, absolute.behavior());
    assertEquals("Queue preview", picture.description());
    assertEquals(ExcelDrawingShapeKind.SIMPLE_SHAPE, shape.kind());
    assertEquals("payload.txt", embeddedObject.fileName());
    assertEquals("OpsPicture.png", picturePayload.fileName());
    assertEquals("Payload", embeddedPayload.label());
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, 0, 0, -1));
    assertThrows(
        IllegalArgumentException.class, () -> new DrawingAnchorReport.OneCell(from, 0L, 20L, null));
    assertThrows(
        IllegalArgumentException.class, () -> new DrawingAnchorReport.OneCell(from, 10L, 0L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(-1L, 2L, 10L, 20L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(1L, -2L, 10L, 20L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(1L, 2L, 0L, 20L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(1L, 2L, 10L, 0L, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Picture(
                " ",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                68L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Picture(
                "OpsPicture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                0L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Picture(
                "OpsPicture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                68L,
                "abc123",
                -1,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Picture(
                "OpsPicture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                68L,
                "abc123",
                null,
                -1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Picture(
                "OpsPicture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                68L,
                "abc123",
                null,
                null,
                " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "OpsShape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, " ", null, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "OpsShape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, "rect", " ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "OpsShape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, "rect", null, -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                " ",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                7L,
                "def456",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                " ",
                "payload.txt",
                "application/octet-stream",
                7L,
                "def456",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                " ",
                "application/octet-stream",
                7L,
                "def456",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                0L,
                "def456",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                7L,
                "def456",
                null,
                4L,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                7L,
                "def456",
                ExcelPictureFormat.PNG,
                4L,
                " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                7L,
                "def456",
                null,
                null,
                "abc123"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "OpsEmbed",
                absolute,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                7L,
                "def456",
                ExcelPictureFormat.PNG,
                0L,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                " ",
                ExcelPictureFormat.PNG,
                "image/png",
                "OpsPicture.png",
                "abc123",
                "cGljdHVyZQ==",
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                "OpsPicture",
                ExcelPictureFormat.PNG,
                "image/png",
                "OpsPicture.png",
                "abc123",
                "not-base64",
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                "OpsPicture",
                ExcelPictureFormat.PNG,
                "image/png",
                "OpsPicture.png",
                "abc123",
                "cGljdHVyZQ==",
                " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.EmbeddedObject(
                "OpsEmbed",
                ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                "application/octet-stream",
                " ",
                "def456",
                "cGF5bG9hZA==",
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.EmbeddedObject(
                "OpsEmbed",
                ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                "application/octet-stream",
                "payload.txt",
                "def456",
                "cGF5bG9hZA==",
                " ",
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.EmbeddedObject(
                "OpsEmbed",
                ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                "application/octet-stream",
                "payload.txt",
                "def456",
                "cGF5bG9hZA==",
                null,
                " "));
  }

  @Test
  void chartReportsValidateReadShapes() {
    DrawingMarkerReport from = new DrawingMarkerReport(1, 2, 0, 0);
    DrawingMarkerReport to = new DrawingMarkerReport(6, 12, 0, 0);
    DrawingAnchorReport.TwoCell anchor =
        new DrawingAnchorReport.TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    ChartReport.Axis categoryAxis =
        new ChartReport.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true);
    ChartReport.Axis valueAxis =
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true);
    ChartReport.Series series =
        new ChartReport.Series(
            new ChartReport.Title.Formula("Chart!$B$1", "Plan"),
            new ChartReport.DataSource.StringReference(
                "ChartCategories", List.of("Jan", "Feb", "Mar")),
            new ChartReport.DataSource.NumericReference(
                "ChartValues", null, List.of("10.0", "18.0", "15.0")));
    ChartReport.Bar bar =
        new ChartReport.Bar(
            "OpsChart",
            anchor,
            new ChartReport.Title.Text("Roadmap"),
            new ChartReport.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            true,
            ExcelChartBarDirection.COLUMN,
            List.of(categoryAxis, valueAxis),
            List.of(series));
    ChartReport.Line line =
        new ChartReport.Line(
            "TrendChart",
            anchor,
            new ChartReport.Title.None(),
            new ChartReport.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            List.of(categoryAxis, valueAxis),
            List.of(
                new ChartReport.Series(
                    new ChartReport.Title.None(),
                    new ChartReport.DataSource.StringLiteral(List.of("Jan", "Feb")),
                    new ChartReport.DataSource.NumericLiteral("0.0", List.of("10", "18")))));
    ChartReport.Pie pie =
        new ChartReport.Pie(
            "Pie",
            anchor,
            new ChartReport.Title.None(),
            new ChartReport.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            180,
            List.of(series));
    ChartReport.Unsupported unsupported =
        new ChartReport.Unsupported(
            "ComboChart",
            anchor,
            List.of("BAR", "LINE"),
            "Only single-plot simple charts are modeled authoritatively.");
    DrawingObjectReport.Chart drawingChart =
        new DrawingObjectReport.Chart("OpsChart", anchor, true, List.of("BAR"), "Roadmap");

    assertEquals("Roadmap", ((ChartReport.Title.Text) bar.title()).text());
    assertEquals(2, bar.axes().size());
    assertEquals(2, line.axes().size());
    assertEquals(180, pie.firstSliceAngle());
    assertEquals(
        "ChartValues",
        ((ChartReport.DataSource.NumericReference) bar.series().getFirst().values()).formula());
    assertEquals(
        List.of("Jan", "Feb"),
        ((ChartReport.DataSource.StringLiteral) line.series().getFirst().categories()).values());
    assertEquals(
        List.of("10", "18"),
        ((ChartReport.DataSource.NumericLiteral) line.series().getFirst().values()).values());
    assertEquals(List.of("BAR", "LINE"), unsupported.plotTypeTokens());
    assertTrue(drawingChart.supported());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Bar(
                " ",
                anchor,
                new ChartReport.Title.Text("Roadmap"),
                new ChartReport.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                ExcelChartDisplayBlanksAs.SPAN,
                false,
                true,
                ExcelChartBarDirection.COLUMN,
                List.of(categoryAxis),
                List.of(series)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartReport.Axis(
                null, ExcelChartAxisPosition.BOTTOM, ExcelChartAxisCrosses.AUTO_ZERO, true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Pie(
                "Pie",
                anchor,
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
                361,
                List.of(series)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartReport.DataSource.StringReference(" ", List.of("Jan")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartReport.DataSource.NumericReference("ChartValues", " ", List.of("10.0")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartReport.DataSource.NumericLiteral(" ", List.of("10.0")));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartReport.Line(
                "Trend",
                anchor,
                null,
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
                List.of(categoryAxis),
                List.of(series)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingObjectReport.Chart("OpsChart", anchor, true, List.of(" "), "Roadmap"));
  }

  @Test
  void chartReportCollectionsDefensivelyCopyAndRejectEmptySeries() {
    DrawingAnchorReport.TwoCell anchor =
        new DrawingAnchorReport.TwoCell(
            new DrawingMarkerReport(1, 2, 0, 0),
            new DrawingMarkerReport(6, 12, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    ChartReport.Axis categoryAxis =
        new ChartReport.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true);
    ChartReport.Axis valueAxis =
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.MIN,
            false);
    ChartReport.Series lineSeries =
        new ChartReport.Series(
            new ChartReport.Title.None(),
            new ChartReport.DataSource.StringLiteral(List.of("Jan", "Feb")),
            new ChartReport.DataSource.NumericLiteral(null, List.of("10", "18")));
    List<ChartReport.Axis> axes = new ArrayList<>(List.of(categoryAxis, valueAxis));
    List<ChartReport.Series> lineSeriesValues = new ArrayList<>(List.of(lineSeries));
    ChartReport.Line copiedLine =
        new ChartReport.Line(
            "CopiedLine",
            anchor,
            new ChartReport.Title.Text("Trend"),
            new ChartReport.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            axes,
            lineSeriesValues);
    List<String> labels = new ArrayList<>(List.of("Jan", "Feb"));
    ChartReport.DataSource.StringLiteral copiedLiteral =
        new ChartReport.DataSource.StringLiteral(labels);
    ChartReport.DataSource.NumericReference formattedReference =
        new ChartReport.DataSource.NumericReference("ChartValues", "0.0", List.of("10.0"));
    ChartReport.Pie pieWithoutAngle =
        new ChartReport.Pie(
            "CopiedPie",
            anchor,
            new ChartReport.Title.None(),
            new ChartReport.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            null,
            List.of(lineSeries));

    axes.clear();
    lineSeriesValues.clear();
    labels.clear();

    assertEquals(2, copiedLine.axes().size());
    assertEquals(1, copiedLine.series().size());
    assertEquals(List.of("Jan", "Feb"), copiedLiteral.values());
    assertEquals("0.0", formattedReference.formatCode());
    assertNull(pieWithoutAngle.firstSliceAngle());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Line(
                "EmptyLine",
                anchor,
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
                List.of(categoryAxis),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Pie(
                "EmptyPie",
                anchor,
                new ChartReport.Title.None(),
                new ChartReport.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
                null,
                List.of()));
  }

  @Test
  void tableConditionalFormattingAndWorkbookProtectionReadContractsValidate() {
    TableEntryReport normalized =
        new TableEntryReport(
            "BudgetTable",
            "Budget",
            "A1:B5",
            1,
            1,
            List.of("Item", "Amount"),
            List.of(
                new TableColumnReport(1L, "Item", null, null, null, null),
                new TableColumnReport(2L, "Amount", "UniqueAmount", "Total", "sum", "[@Amount]*2")),
            new TableStyleReport.Named("TableStyleMedium2", false, false, true, false),
            true,
            null,
            true,
            true,
            false,
            null,
            null,
            null);

    assertEquals("", normalized.comment());
    assertEquals("", normalized.headerRowCellStyle());
    assertEquals("UniqueAmount", normalized.columns().get(1).uniqueName());
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableColumnReport(-1L, "Item", null, null, null, null));

    DifferentialStyleReport style =
        new DifferentialStyleReport(
            "0.00", true, null, null, "#AABBCC", null, null, null, null, List.of());
    ConditionalFormattingRuleReport.Top10Rule top10 =
        new ConditionalFormattingRuleReport.Top10Rule(1, false, 10, true, false, style);

    assertEquals(10, top10.rank());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleReport.Top10Rule(1, false, -1, false, false, style));

    WorkbookReadOperation.GetWorkbookProtection read =
        new WorkbookReadOperation.GetWorkbookProtection("workbook-protection");
    WorkbookReadResult.WorkbookProtectionResult result =
        new WorkbookReadResult.WorkbookProtectionResult(
            "workbook-protection", new WorkbookProtectionReport(true, false, true, true, false));

    assertEquals("workbook-protection", read.requestId());
    assertTrue(result.protection().workbookPasswordHashPresent());
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadOperation.GetWorkbookProtection(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadOperation.GetWorkbookProtection(" "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.WorkbookProtectionResult("workbook-protection", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookProtectionResult(
                " ", new WorkbookProtectionReport(true, false, true, true, false)));
  }

  private static CellColorReport rgb(String rgb) {
    return new CellColorReport(rgb);
  }

  private static CellFontReport font(CellColorReport color) {
    return new CellFontReport(
        false,
        false,
        "Aptos",
        new FontHeightReport(220, BigDecimal.valueOf(11)),
        color,
        false,
        false);
  }
}
