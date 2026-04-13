package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.math.BigDecimal;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for converting advanced engine read results into protocol response shapes. */
class WorkbookReadResultConverterTest {
  @Test
  void convertsPackageSecurityReadResultsIntoProtocolShapes() {
    WorkbookReadResult.PackageSecurityResult packageSecurity =
        assertInstanceOf(
            WorkbookReadResult.PackageSecurityResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PackageSecurityResult(
                    "security",
                    new ExcelOoxmlPackageSecuritySnapshot(
                        new ExcelOoxmlEncryptionSnapshot(
                            true,
                            ExcelOoxmlEncryptionMode.AGILE,
                            "AES256",
                            "SHA512",
                            "CBC",
                            256,
                            16,
                            100_000),
                        List.of(
                            new ExcelOoxmlSignatureSnapshot(
                                "/_xmlsignatures/sig1.xml",
                                "CN=GridGrind Signing Test",
                                "CN=GridGrind Signing Test",
                                "01AB",
                                ExcelOoxmlSignatureState.VALID))))));

    assertTrue(packageSecurity.security().encryption().encrypted());
    assertEquals(
        ExcelOoxmlSignatureState.VALID, packageSecurity.security().signatures().getFirst().state());
  }

  @Test
  void convertsPlainCommentAndWorkbookProtectionFactsDirectly() {
    assertNull(WorkbookReadResultConverter.toCommentReport((ExcelComment) null));
    GridGrindResponse.CommentReport plainComment =
        WorkbookReadResultConverter.toCommentReport(new ExcelComment("Review", "GridGrind", false));
    WorkbookProtectionReport protection =
        WorkbookReadResultConverter.toWorkbookProtectionReport(
            new ExcelWorkbookProtectionSnapshot(true, false, true, true, false));

    assertEquals("Review", plainComment.text());
    assertEquals("GridGrind", plainComment.author());
    assertNull(plainComment.runs());
    assertTrue(protection.structureLocked());
    assertTrue(protection.revisionLocked());
  }

  @Test
  void convertsAdvancedReadResultsIntoProtocolShapes() {
    WorkbookReadResult.WorkbookProtectionResult protection =
        assertInstanceOf(
            WorkbookReadResult.WorkbookProtectionResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookProtectionResult(
                    "workbook-protection",
                    new ExcelWorkbookProtectionSnapshot(true, false, true, true, false))));
    assertTrue(protection.protection().structureLocked());

    WorkbookReadResult.CellsResult cells =
        assertInstanceOf(
            WorkbookReadResult.CellsResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult(
                    "cells", "Budget", List.of(advancedCell()))));
    GridGrindResponse.CellReport.TextReport cell =
        assertInstanceOf(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst());
    assertEquals(new CellColorReport(null, 2, null, 0.25d), cell.style().font().fontColor());
    assertEquals(new CellColorReport(null, null, 12, null), cell.style().border().bottom().color());
    assertNotNull(cell.style().fill().gradient());
    assertEquals(2, cell.style().fill().gradient().stops().size());
    assertEquals("https://example.com/tasks/42", ((HyperlinkTarget.Url) cell.hyperlink()).target());
    assertNotNull(cell.comment());
    assertEquals(2, cell.comment().runs().size());
    assertEquals(1, cell.comment().anchor().firstColumn());

    WorkbookReadResult.CommentsResult comments =
        assertInstanceOf(
            WorkbookReadResult.CommentsResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.CellComment(
                            "C3", richComment())))));
    assertEquals(2, comments.comments().getFirst().comment().runs().size());
    assertEquals(6, comments.comments().getFirst().comment().anchor().lastRow());

    WorkbookReadResult.PrintLayoutResult printLayout =
        assertInstanceOf(
            WorkbookReadResult.PrintLayoutResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                    "print-layout", "Budget", advancedPrintLayout())));
    assertEquals(List.of(10, 20), printLayout.layout().setup().rowBreaks());
    assertEquals(9, printLayout.layout().setup().paperSize());

    WorkbookReadResult.AutofiltersResult autofilters =
        assertInstanceOf(
            WorkbookReadResult.AutofiltersResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult(
                    "autofilters", "Budget", advancedAutofilters())));
    AutofilterEntryReport.SheetOwned sheetOwned =
        assertInstanceOf(
            AutofilterEntryReport.SheetOwned.class, autofilters.autofilters().getFirst());
    assertEquals(6, sheetOwned.filterColumns().size());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Custom.class,
        sheetOwned.filterColumns().get(1).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Dynamic.class,
        sheetOwned.filterColumns().get(2).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Top10.class, sheetOwned.filterColumns().get(3).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Color.class, sheetOwned.filterColumns().get(4).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Icon.class, sheetOwned.filterColumns().get(5).criterion());
    assertEquals("A1:F5", sheetOwned.sortState().range());
    assertInstanceOf(AutofilterEntryReport.TableOwned.class, autofilters.autofilters().get(1));

    WorkbookReadResult.TablesResult tables =
        assertInstanceOf(
            WorkbookReadResult.TablesResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult(
                    "tables", List.of(advancedTable()))));
    assertEquals("HeaderStyle", tables.tables().getFirst().headerRowCellStyle());
    assertEquals("Total", tables.tables().getFirst().columns().get(1).totalsRowLabel());

    WorkbookReadResult.ConditionalFormattingResult conditionalFormatting =
        assertInstanceOf(
            WorkbookReadResult.ConditionalFormattingResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult(
                    "conditional-formatting",
                    "Budget",
                    List.of(
                        new ExcelConditionalFormattingBlockSnapshot(
                            List.of("A1:A5"),
                            List.of(
                                new ExcelConditionalFormattingRuleSnapshot.Top10Rule(
                                    1, false, 10, true, false, differentialStyle())))))));
    assertInstanceOf(
        ConditionalFormattingRuleReport.Top10Rule.class,
        conditionalFormatting.conditionalFormattingBlocks().getFirst().rules().getFirst());

    WorkbookReadResult.DrawingObjectsResult drawingObjects =
        assertInstanceOf(
            WorkbookReadResult.DrawingObjectsResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Budget",
                    List.of(
                        new ExcelDrawingObjectSnapshot.Picture(
                            "OpsPicture",
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(1, 2, 3, 4),
                                new ExcelDrawingMarker(4, 6, 7, 8),
                                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                            ExcelPictureFormat.PNG,
                            "image/png",
                            68L,
                            "abc123",
                            null,
                            null,
                            "Queue preview"),
                        new ExcelDrawingObjectSnapshot.Shape(
                            "OpsShape",
                            new ExcelDrawingAnchor.OneCell(
                                new ExcelDrawingMarker(5, 6, 0, 0), 10L, 20L, null),
                            ExcelDrawingShapeKind.SIMPLE_SHAPE,
                            "rect",
                            "Queue",
                            0),
                        new ExcelDrawingObjectSnapshot.Chart(
                            "OpsChart",
                            new ExcelDrawingAnchor.TwoCell(
                                new ExcelDrawingMarker(1, 2, 0, 0),
                                new ExcelDrawingMarker(6, 12, 0, 0),
                                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                            true,
                            List.of("BAR"),
                            "Roadmap"),
                        new ExcelDrawingObjectSnapshot.EmbeddedObject(
                            "OpsEmbed",
                            new ExcelDrawingAnchor.Absolute(1L, 2L, 10L, 20L, null),
                            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                            "Payload",
                            "payload.txt",
                            "payload.txt",
                            "application/octet-stream",
                            7L,
                            "def456",
                            null,
                            null,
                            null)))));
    WorkbookReadResult.DrawingObjectPayloadResult drawingPayload =
        assertInstanceOf(
            WorkbookReadResult.DrawingObjectPayloadResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.DrawingObjectPayloadResult(
                    "drawing-payload",
                    "Budget",
                    new ExcelDrawingObjectPayload.EmbeddedObject(
                        "OpsEmbed",
                        ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                        "application/octet-stream",
                        "payload.txt",
                        "def456",
                        new ExcelBinaryData(
                            "payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                        "Payload",
                        "payload.txt"))));
    assertEquals(4, drawingObjects.drawingObjects().size());
    DrawingObjectReport.Picture picture =
        assertInstanceOf(DrawingObjectReport.Picture.class, drawingObjects.drawingObjects().get(0));
    assertEquals(ExcelPictureFormat.PNG, picture.format());
    assertEquals(
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE,
        assertInstanceOf(DrawingAnchorReport.TwoCell.class, picture.anchor()).behavior());
    assertInstanceOf(DrawingObjectReport.Shape.class, drawingObjects.drawingObjects().get(1));
    DrawingObjectReport.Chart chartObject =
        assertInstanceOf(DrawingObjectReport.Chart.class, drawingObjects.drawingObjects().get(2));
    assertTrue(chartObject.supported());
    assertEquals(List.of("BAR"), chartObject.plotTypeTokens());
    assertInstanceOf(
        DrawingObjectReport.EmbeddedObject.class, drawingObjects.drawingObjects().get(3));
    assertEquals("cGF5bG9hZA==", drawingPayload.payload().base64Data());
    DrawingObjectPayloadReport.Picture picturePayload =
        assertInstanceOf(
            DrawingObjectPayloadReport.Picture.class,
            WorkbookReadResultConverter.toDrawingObjectPayloadReport(
                new ExcelDrawingObjectPayload.Picture(
                    "OpsPicture",
                    ExcelPictureFormat.PNG,
                    "image/png",
                    "OpsPicture.png",
                    "abc123",
                    new ExcelBinaryData(
                        "picture".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                    "Queue preview")));
    assertEquals("OpsPicture.png", picturePayload.fileName());
    assertEquals("cGljdHVyZQ==", picturePayload.base64Data());
  }

  @Test
  void convertsPivotReadResultsIntoProtocolShapes() {
    WorkbookReadResult.PivotTablesResult pivotTables =
        assertInstanceOf(
            WorkbookReadResult.PivotTablesResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PivotTablesResult(
                    "pivots",
                    List.of(
                        new ExcelPivotTableSnapshot.Supported(
                            "Sales Pivot 2026",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("C5", "C5:G9"),
                            new ExcelPivotTableSnapshot.Source.Table("SalesTable", "Data", "A1:D5"),
                            List.of(new ExcelPivotTableSnapshot.Field(0, "Region")),
                            List.of(new ExcelPivotTableSnapshot.Field(1, "Stage")),
                            List.of(new ExcelPivotTableSnapshot.Field(2, "Owner")),
                            List.of(
                                new ExcelPivotTableSnapshot.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    "#,##0.00")),
                            true),
                        new ExcelPivotTableSnapshot.Unsupported(
                            "Broken Pivot",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("A3", "A3:C8"),
                            "Pivot cache source no longer resolves cleanly.")))));
    WorkbookReadResult.PivotTableHealthResult pivotTableHealth =
        assertInstanceOf(
            WorkbookReadResult.PivotTableHealthResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PivotTableHealthResult(
                    "pivot-health",
                    new WorkbookAnalysis.PivotTableHealth(
                        1,
                        new WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0),
                        List.of(
                            new WorkbookAnalysis.AnalysisFinding(
                                WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
                                WorkbookAnalysis.AnalysisSeverity.WARNING,
                                "Pivot table name is missing",
                                "GridGrind assigned a synthetic identifier for readback.",
                                new WorkbookAnalysis.AnalysisLocation.Sheet("Report"),
                                List.of("_GG_PIVOT_Report_A3")))))));

    PivotTableReport.Supported supported =
        assertInstanceOf(PivotTableReport.Supported.class, pivotTables.pivotTables().getFirst());
    PivotTableReport.Unsupported unsupported =
        assertInstanceOf(PivotTableReport.Unsupported.class, pivotTables.pivotTables().get(1));

    assertEquals("SalesTable", ((PivotTableReport.Source.Table) supported.source()).name());
    assertEquals("Amount", supported.dataFields().getFirst().sourceColumnName());
    assertTrue(supported.valuesAxisOnColumns());
    assertEquals("Pivot cache source no longer resolves cleanly.", unsupported.detail());
    assertEquals(1, pivotTableHealth.analysis().checkedPivotTableCount());
    assertEquals(
        AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
        pivotTableHealth.analysis().findings().getFirst().code());
  }

  @Test
  void convertsPivotSourceVariantsIntoProtocolShapes() {
    WorkbookReadResult.PivotTablesResult pivotTables =
        assertInstanceOf(
            WorkbookReadResult.PivotTablesResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PivotTablesResult(
                    "pivots",
                    List.of(
                        new ExcelPivotTableSnapshot.Supported(
                            "Range Pivot",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("C5", "C5:G9"),
                            new ExcelPivotTableSnapshot.Source.Range("Data", "A1:D5"),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new ExcelPivotTableSnapshot.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    null)),
                            false),
                        new ExcelPivotTableSnapshot.Supported(
                            "Named Range Pivot",
                            "Report",
                            new ExcelPivotTableSnapshot.Anchor("A3", "A3:E8"),
                            new ExcelPivotTableSnapshot.Source.NamedRange(
                                "PivotSource", "Data", "A1:D5"),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new ExcelPivotTableSnapshot.DataField(
                                    3,
                                    "Amount",
                                    ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    null)),
                            false)))));

    PivotTableReport.Supported rangePivot =
        assertInstanceOf(PivotTableReport.Supported.class, pivotTables.pivotTables().getFirst());
    PivotTableReport.Supported namedRangePivot =
        assertInstanceOf(PivotTableReport.Supported.class, pivotTables.pivotTables().get(1));

    assertEquals("Data", ((PivotTableReport.Source.Range) rangePivot.source()).sheetName());
    assertEquals(
        "PivotSource", ((PivotTableReport.Source.NamedRange) namedRangePivot.source()).name());
  }

  @Test
  void convertsChartReadResultsIntoProtocolShapes() {
    WorkbookReadResult.ChartsResult charts =
        assertInstanceOf(
            WorkbookReadResult.ChartsResult.class,
            WorkbookReadResultConverter.toReadResult(baseChartsResult()));
    ChartReport.Bar chartReport =
        assertInstanceOf(ChartReport.Bar.class, charts.charts().getFirst());
    assertEquals("Roadmap", ((ChartReport.Title.Text) chartReport.title()).text());
    assertEquals(
        "ChartValues",
        ((ChartReport.DataSource.NumericReference) chartReport.series().getFirst().values())
            .formula());

    WorkbookReadResult.ChartsResult advancedCharts =
        assertInstanceOf(
            WorkbookReadResult.ChartsResult.class,
            WorkbookReadResultConverter.toReadResult(advancedChartsResult()));
    ChartReport.Line lineReport =
        assertInstanceOf(ChartReport.Line.class, advancedCharts.charts().get(0));
    assertTrue(lineReport.title() instanceof ChartReport.Title.None);
    assertEquals(
        List.of("Jan", "Feb"),
        ((ChartReport.DataSource.StringLiteral) lineReport.series().getFirst().categories())
            .values());
    assertEquals(
        List.of("10", "18"),
        ((ChartReport.DataSource.NumericLiteral) lineReport.series().getFirst().values()).values());
    ChartReport.Pie pieReport =
        assertInstanceOf(ChartReport.Pie.class, advancedCharts.charts().get(1));
    assertEquals(120, pieReport.firstSliceAngle());
    assertEquals("Actual", ((ChartReport.Title.Formula) pieReport.title()).cachedText());
    ChartReport.Unsupported unsupportedChart =
        assertInstanceOf(ChartReport.Unsupported.class, advancedCharts.charts().get(2));
    assertEquals(List.of("AREA"), unsupportedChart.plotTypeTokens());
  }

  @Test
  void directChartReportConversionCoversStandaloneSwitchBranches() {
    ExcelChartSnapshot lineSnapshot =
        new ExcelChartSnapshot.Line(
            "OpsLine",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(2, 3, 0, 0),
                new ExcelDrawingMarker(7, 14, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            new ExcelChartSnapshot.Title.None(),
            new ExcelChartSnapshot.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            List.of(
                new ExcelChartSnapshot.Axis(
                    ExcelChartAxisKind.CATEGORY,
                    ExcelChartAxisPosition.TOP,
                    ExcelChartAxisCrosses.AUTO_ZERO,
                    true),
                new ExcelChartSnapshot.Axis(
                    ExcelChartAxisKind.VALUE,
                    ExcelChartAxisPosition.RIGHT,
                    ExcelChartAxisCrosses.MAX,
                    false)),
            List.of(
                new ExcelChartSnapshot.Series(
                    new ExcelChartSnapshot.Title.None(),
                    new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan", "Feb")),
                    new ExcelChartSnapshot.DataSource.NumericLiteral("0.0", List.of("10", "18")))));
    ExcelChartSnapshot pieSnapshot =
        new ExcelChartSnapshot.Pie(
            "OpsPie",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(8, 3, 0, 0),
                new ExcelDrawingMarker(13, 14, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            new ExcelChartSnapshot.Title.Formula("Budget!$C$1", "Actual"),
            new ExcelChartSnapshot.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            120,
            List.of(
                new ExcelChartSnapshot.Series(
                    new ExcelChartSnapshot.Title.Text("Actual"),
                    new ExcelChartSnapshot.DataSource.StringReference(
                        "ChartCategories", List.of("Jan", "Feb")),
                    new ExcelChartSnapshot.DataSource.NumericReference(
                        "ChartActual", "0.0", List.of("12", "16")))));
    ExcelChartSnapshot unsupportedSnapshot =
        new ExcelChartSnapshot.Unsupported(
            "AreaOnly",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(14, 3, 0, 0),
                new ExcelDrawingMarker(19, 14, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            List.of("AREA"),
            "Only simple single-plot charts are modeled.");

    ChartReport.Line lineReport =
        assertInstanceOf(
            ChartReport.Line.class,
            WorkbookReadResultConverter.toChartReport((ExcelChartSnapshot) lineSnapshot));
    ChartReport.Pie pieReport =
        assertInstanceOf(
            ChartReport.Pie.class,
            WorkbookReadResultConverter.toChartReport((ExcelChartSnapshot) pieSnapshot));
    ChartReport.Unsupported unsupportedReport =
        assertInstanceOf(
            ChartReport.Unsupported.class,
            WorkbookReadResultConverter.toChartReport((ExcelChartSnapshot) unsupportedSnapshot));

    assertTrue(lineReport.title() instanceof ChartReport.Title.None);
    assertTrue(lineReport.legend() instanceof ChartReport.Legend.Hidden);
    assertEquals(
        List.of("Jan", "Feb"),
        assertInstanceOf(
                ChartReport.DataSource.StringLiteral.class,
                lineReport.series().getFirst().categories())
            .values());
    assertEquals(
        "0.0",
        assertInstanceOf(
                ChartReport.DataSource.NumericLiteral.class,
                lineReport.series().getFirst().values())
            .formatCode());
    assertEquals(
        "Actual",
        assertInstanceOf(ChartReport.Title.Formula.class, pieReport.title()).cachedText());
    assertEquals(
        List.of("12", "16"),
        assertInstanceOf(
                ChartReport.DataSource.NumericReference.class,
                pieReport.series().getFirst().values())
            .cachedValues());
    assertEquals(List.of("AREA"), unsupportedReport.plotTypeTokens());
  }

  private static dev.erst.gridgrind.excel.WorkbookReadResult.ChartsResult baseChartsResult() {
    return new dev.erst.gridgrind.excel.WorkbookReadResult.ChartsResult(
        "charts",
        "Budget",
        List.of(
            new ExcelChartSnapshot.Bar(
                "OpsChart",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(1, 2, 0, 0),
                    new ExcelDrawingMarker(6, 12, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                new ExcelChartSnapshot.Title.Text("Roadmap"),
                new ExcelChartSnapshot.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                ExcelChartDisplayBlanksAs.SPAN,
                false,
                true,
                ExcelChartBarDirection.COLUMN,
                List.of(
                    new ExcelChartSnapshot.Axis(
                        ExcelChartAxisKind.CATEGORY,
                        ExcelChartAxisPosition.BOTTOM,
                        ExcelChartAxisCrosses.AUTO_ZERO,
                        true),
                    new ExcelChartSnapshot.Axis(
                        ExcelChartAxisKind.VALUE,
                        ExcelChartAxisPosition.LEFT,
                        ExcelChartAxisCrosses.AUTO_ZERO,
                        true)),
                List.of(
                    new ExcelChartSnapshot.Series(
                        new ExcelChartSnapshot.Title.Formula("Chart!$B$1", "Plan"),
                        new ExcelChartSnapshot.DataSource.StringReference(
                            "ChartCategories", List.of("Jan", "Feb", "Mar")),
                        new ExcelChartSnapshot.DataSource.NumericReference(
                            "ChartValues", null, List.of("10.0", "18.0", "15.0")))))));
  }

  private static dev.erst.gridgrind.excel.WorkbookReadResult.ChartsResult advancedChartsResult() {
    return new dev.erst.gridgrind.excel.WorkbookReadResult.ChartsResult(
        "charts-advanced",
        "Budget",
        List.of(
            new ExcelChartSnapshot.Line(
                "TrendChart",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(2, 3, 0, 0),
                    new ExcelDrawingMarker(7, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
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
                List.of(
                    new ExcelChartSnapshot.Series(
                        new ExcelChartSnapshot.Title.None(),
                        new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan", "Feb")),
                        new ExcelChartSnapshot.DataSource.NumericLiteral(
                            "0.0", List.of("10", "18"))))),
            new ExcelChartSnapshot.Pie(
                "ShareChart",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(8, 3, 0, 0),
                    new ExcelDrawingMarker(13, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                new ExcelChartSnapshot.Title.Formula("Budget!$C$1", "Actual"),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.ZERO,
                false,
                true,
                120,
                List.of(
                    new ExcelChartSnapshot.Series(
                        new ExcelChartSnapshot.Title.Text("Actual"),
                        new ExcelChartSnapshot.DataSource.StringReference(
                            "ChartCategories", List.of("Jan", "Feb")),
                        new ExcelChartSnapshot.DataSource.NumericReference(
                            "ChartActual", "0.0", List.of("12", "16"))))),
            new ExcelChartSnapshot.Unsupported(
                "AreaOnly",
                new ExcelDrawingAnchor.TwoCell(
                    new ExcelDrawingMarker(14, 3, 0, 0),
                    new ExcelDrawingMarker(19, 14, 0, 0),
                    ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                List.of("AREA"),
                "Chart plot family is outside the current modeled simple-chart contract.")));
  }

  private static ExcelCellSnapshot advancedCell() {
    ExcelRichTextSnapshot richText =
        new ExcelRichTextSnapshot(List.of(new ExcelRichTextRunSnapshot("Styled", advancedFont())));
    return new ExcelCellSnapshot.TextSnapshot(
        "C3",
        "STRING",
        "Styled",
        advancedStyle(),
        ExcelCellMetadataSnapshot.of(
            new ExcelHyperlink.Url("https://example.com/tasks/42"), richComment()),
        "Styled",
        richText);
  }

  private static ExcelCommentSnapshot richComment() {
    return new ExcelCommentSnapshot(
        "Hi there",
        "Ada",
        true,
        new ExcelRichTextSnapshot(
            List.of(
                new ExcelRichTextRunSnapshot("Hi ", advancedFont()),
                new ExcelRichTextRunSnapshot("there", accentFont()))),
        new ExcelCommentAnchorSnapshot(1, 2, 4, 6));
  }

  private static ExcelCellStyleSnapshot advancedStyle() {
    return new ExcelCellStyleSnapshot(
        "0.00",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.CENTER, 0, 0),
        advancedFont(),
        new ExcelCellFillSnapshot(
            ExcelFillPattern.NONE,
            null,
            null,
            new ExcelGradientFillSnapshot(
                "LINEAR",
                45.0d,
                null,
                null,
                null,
                null,
                List.of(
                    new ExcelGradientStopSnapshot(0.0d, new ExcelColorSnapshot("#112233")),
                    new ExcelGradientStopSnapshot(
                        1.0d, new ExcelColorSnapshot(null, 4, null, 0.45d))))),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(
                ExcelBorderStyle.THICK, new ExcelColorSnapshot(null, null, 12, null)),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  private static ExcelCellFontSnapshot advancedFont() {
    return new ExcelCellFontSnapshot(
        true,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        new ExcelColorSnapshot(null, 2, null, 0.25d),
        false,
        false);
  }

  private static ExcelCellFontSnapshot accentFont() {
    return new ExcelCellFontSnapshot(
        false,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        new ExcelColorSnapshot("#AABBCC"),
        false,
        false);
  }

  private static ExcelPrintLayoutSnapshot advancedPrintLayout() {
    return new ExcelPrintLayoutSnapshot(
        new ExcelPrintLayout(
            new ExcelPrintLayout.Area.Range("A1:F20"),
            ExcelPrintOrientation.LANDSCAPE,
            new ExcelPrintLayout.Scaling.Fit(1, 0),
            new ExcelPrintLayout.TitleRows.Band(0, 0),
            new ExcelPrintLayout.TitleColumns.None(),
            new ExcelHeaderFooterText("Ops", "Queue", ""),
            new ExcelHeaderFooterText("", "Internal", "Page &P")),
        new ExcelPrintSetupSnapshot(
            new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
            true,
            false,
            9,
            false,
            true,
            2,
            true,
            3,
            List.of(10, 20),
            List.of(2, 4)));
  }

  private static List<ExcelAutofilterSnapshot> advancedAutofilters() {
    ExcelAutofilterSortStateSnapshot sortState =
        new ExcelAutofilterSortStateSnapshot(
            "A1:F5",
            true,
            false,
            "",
            List.of(
                new ExcelAutofilterSortConditionSnapshot(
                    "A2:A5", true, "", new ExcelColorSnapshot("#AABBCC"), 1)));
    List<ExcelAutofilterFilterColumnSnapshot> filterColumns =
        List.of(
            new ExcelAutofilterFilterColumnSnapshot(
                0L,
                false,
                new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Queued"), true)),
            new ExcelAutofilterFilterColumnSnapshot(
                1L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Custom(
                    true,
                    List.of(
                        new ExcelAutofilterFilterCriterionSnapshot.CustomCondition(
                            "equal", "Ada")))),
            new ExcelAutofilterFilterColumnSnapshot(
                2L, true, new ExcelAutofilterFilterCriterionSnapshot.Dynamic("TODAY", 1.0d, 2.0d)),
            new ExcelAutofilterFilterColumnSnapshot(
                3L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, 10.0d, 8.0d)),
            new ExcelAutofilterFilterColumnSnapshot(
                4L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Color(
                    false, new ExcelColorSnapshot(null, 4, null, 0.45d))),
            new ExcelAutofilterFilterColumnSnapshot(
                5L, true, new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", 2)));
    return List.of(
        new ExcelAutofilterSnapshot.SheetOwned("A1:F5", filterColumns, sortState),
        new ExcelAutofilterSnapshot.TableOwned("H1:I5", "QueueTable", List.of(), sortState));
  }

  private static ExcelTableSnapshot advancedTable() {
    return new ExcelTableSnapshot(
        "QueueTable",
        "Budget",
        "A1:B5",
        1,
        1,
        List.of("Item", "Amount"),
        List.of(
            new ExcelTableColumnSnapshot(1L, "Item", "", "", "", ""),
            new ExcelTableColumnSnapshot(
                2L, "Amount", "UniqueAmount", "Total", "sum", "[@Amount]*2")),
        new ExcelTableStyleSnapshot.Named("TableStyleMedium2", false, false, true, false),
        true,
        "Queue comment",
        true,
        true,
        false,
        "HeaderStyle",
        "DataStyle",
        "TotalsStyle");
  }

  private static ExcelDifferentialStyleSnapshot differentialStyle() {
    return new ExcelDifferentialStyleSnapshot(
        "0.00", true, null, null, "#AABBCC", null, null, null, null, List.of());
  }
}
