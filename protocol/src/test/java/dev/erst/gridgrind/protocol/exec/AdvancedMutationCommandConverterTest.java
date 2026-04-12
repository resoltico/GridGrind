package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterion;
import dev.erst.gridgrind.excel.ExcelAutofilterSortCondition;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentAnchor;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.protocol.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.protocol.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.protocol.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.protocol.dto.CellBorderInput;
import dev.erst.gridgrind.protocol.dto.CellBorderSideInput;
import dev.erst.gridgrind.protocol.dto.CellFillInput;
import dev.erst.gridgrind.protocol.dto.CellFontInput;
import dev.erst.gridgrind.protocol.dto.CellGradientFillInput;
import dev.erst.gridgrind.protocol.dto.CellGradientStopInput;
import dev.erst.gridgrind.protocol.dto.CellStyleInput;
import dev.erst.gridgrind.protocol.dto.ChartInput;
import dev.erst.gridgrind.protocol.dto.ColorInput;
import dev.erst.gridgrind.protocol.dto.CommentAnchorInput;
import dev.erst.gridgrind.protocol.dto.CommentInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.protocol.dto.DrawingAnchorInput;
import dev.erst.gridgrind.protocol.dto.DrawingMarkerInput;
import dev.erst.gridgrind.protocol.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.protocol.dto.NamedRangeScope;
import dev.erst.gridgrind.protocol.dto.NamedRangeTarget;
import dev.erst.gridgrind.protocol.dto.PictureDataInput;
import dev.erst.gridgrind.protocol.dto.PictureInput;
import dev.erst.gridgrind.protocol.dto.RichTextRunInput;
import dev.erst.gridgrind.protocol.dto.ShapeInput;
import dev.erst.gridgrind.protocol.dto.SheetProtectionSettings;
import dev.erst.gridgrind.protocol.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused command-converter coverage for advanced workbook-core mutation payloads. */
class AdvancedMutationCommandConverterTest {
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  @Test
  void convertsAdvancedCommentProtectionAutofilterAndNamedRangeOperations() {
    WorkbookCommand.SetComment commentCommand =
        assertInstanceOf(
            WorkbookCommand.SetComment.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetComment(
                    "Budget",
                    "B4",
                    new CommentInput(
                        "Ada Lovelace",
                        "GridGrind",
                        true,
                        List.of(
                            new RichTextRunInput("Ada", null),
                            new RichTextRunInput(" Lovelace", null)),
                        new CommentAnchorInput(1, 2, 4, 6)))));
    assertEquals(
        new ExcelComment(
            "Ada Lovelace",
            "GridGrind",
            true,
            new ExcelRichText(
                List.of(
                    new ExcelRichTextRun("Ada", null), new ExcelRichTextRun(" Lovelace", null))),
            new ExcelCommentAnchor(1, 2, 4, 6)),
        commentCommand.comment());

    WorkbookCommand.SetWorkbookProtection protectionCommand =
        assertInstanceOf(
            WorkbookCommand.SetWorkbookProtection.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetWorkbookProtection(
                    new WorkbookProtectionInput(
                        null, true, null, "book-secret", "review-secret"))));
    assertEquals(
        new ExcelWorkbookProtectionSettings(false, true, false, "book-secret", "review-secret"),
        protectionCommand.protection());

    assertInstanceOf(
        WorkbookCommand.ClearWorkbookProtection.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.ClearWorkbookProtection()));

    WorkbookCommand.SetAutofilter simpleAutofilter =
        assertInstanceOf(
            WorkbookCommand.SetAutofilter.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetAutofilter("Budget", "A1:C4")));
    assertEquals(List.of(), simpleAutofilter.criteria());
    assertNull(simpleAutofilter.sortState());
    assertEquals(
        List.of(), new WorkbookOperation.SetAutofilter("Budget", "A1:C4", null, null).criteria());

    WorkbookCommand.SetAutofilter advancedAutofilter =
        assertInstanceOf(
            WorkbookCommand.SetAutofilter.class,
            WorkbookCommandConverter.toCommand(advancedAutofilterOperation()));
    assertEquals(expectedAutofilterCriteria(), advancedAutofilter.criteria());
    assertEquals(expectedAutofilterSortState(), advancedAutofilter.sortState());

    WorkbookCommand.SetNamedRange namedRangeCommand =
        assertInstanceOf(
            WorkbookCommand.SetNamedRange.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetNamedRange(
                    "BudgetExpr",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("SUM(Budget!A1:A3)"))));
    assertEquals(
        new ExcelNamedRangeDefinition(
            "BudgetExpr",
            new ExcelNamedRangeScope.WorkbookScope(),
            new ExcelNamedRangeTarget("SUM(Budget!A1:A3)")),
        namedRangeCommand.definition());

    WorkbookOperation.SetSheetProtection sheetProtection =
        new WorkbookOperation.SetSheetProtection(
            "Budget",
            new SheetProtectionSettings(
                true, false, false, false, false, false, false, false, false, false, false, false,
                false, false, false),
            "sheet-secret");
    assertEquals("sheet-secret", sheetProtection.password());
    org.junit.jupiter.api.Assertions.assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetSheetProtection(
                "Budget",
                new SheetProtectionSettings(
                    true, false, false, false, false, false, false, false, false, false, false,
                    false, false, false, false),
                " "));
  }

  @Test
  void convertsDrawingMutationOperations() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 3, 4),
            new DrawingMarkerInput(4, 6, 7, 8),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    PictureDataInput pictureData = new PictureDataInput(ExcelPictureFormat.PNG, PNG_PIXEL_BASE64);

    WorkbookCommand.SetPicture pictureCommand =
        assertInstanceOf(
            WorkbookCommand.SetPicture.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetPicture(
                    "Ops", new PictureInput("OpsPicture", pictureData, anchor, "Queue preview"))));
    WorkbookCommand.SetShape shapeCommand =
        assertInstanceOf(
            WorkbookCommand.SetShape.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetShape(
                    "Ops",
                    new ShapeInput(
                        "OpsShape",
                        ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                        anchor,
                        "rect",
                        "Queue"))));
    WorkbookCommand.SetEmbeddedObject embeddedObjectCommand =
        assertInstanceOf(
            WorkbookCommand.SetEmbeddedObject.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetEmbeddedObject(
                    "Ops",
                    new EmbeddedObjectInput(
                        "OpsEmbed",
                        "Payload",
                        "payload.txt",
                        "payload.txt",
                        "cGF5bG9hZA==",
                        pictureData,
                        anchor))));
    WorkbookCommand.SetDrawingObjectAnchor moveCommand =
        assertInstanceOf(
            WorkbookCommand.SetDrawingObjectAnchor.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetDrawingObjectAnchor("Ops", "OpsPicture", anchor)));
    WorkbookCommand.DeleteDrawingObject deleteCommand =
        assertInstanceOf(
            WorkbookCommand.DeleteDrawingObject.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.DeleteDrawingObject("Ops", "OpsPicture")));

    assertEquals(
        new ExcelPictureDefinition(
            "OpsPicture",
            new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
            ExcelPictureFormat.PNG,
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 4),
                new ExcelDrawingMarker(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
            "Queue preview"),
        pictureCommand.picture());
    assertEquals(
        new ExcelShapeDefinition(
            "OpsShape",
            ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 4),
                new ExcelDrawingMarker(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
            "rect",
            "Queue"),
        shapeCommand.shape());
    assertEquals(
        new ExcelEmbeddedObjectDefinition(
            "OpsEmbed",
            "Payload",
            "payload.txt",
            "payload.txt",
            new ExcelBinaryData("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
            ExcelPictureFormat.PNG,
            new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 4),
                new ExcelDrawingMarker(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE)),
        embeddedObjectCommand.embeddedObject());
    assertEquals("OpsPicture", moveCommand.objectName());
    assertEquals("OpsPicture", deleteCommand.objectName());
  }

  @Test
  void convertsChartMutationOperationsAcrossSupportedFamilies() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 3, 4),
            new DrawingMarkerInput(7, 12, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ChartInput.Series firstSeries =
        new ChartInput.Series(
            new ChartInput.Title.Formula("B1"),
            new ChartInput.DataSource("A2:A4"),
            new ChartInput.DataSource("B2:B4"));
    ChartInput.Series secondSeries =
        new ChartInput.Series(
            new ChartInput.Title.Text("Actual"),
            new ChartInput.DataSource("ChartCategories"),
            new ChartInput.DataSource("ChartActual"));

    WorkbookCommand.SetChart lineCommand =
        assertInstanceOf(
            WorkbookCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetChart(
                    "Ops",
                    new ChartInput.Line(
                        "TrendChart",
                        anchor,
                        new ChartInput.Title.Text("Trend"),
                        new ChartInput.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.ZERO,
                        false,
                        true,
                        List.of(firstSeries)))));
    WorkbookCommand.SetChart pieCommand =
        assertInstanceOf(
            WorkbookCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetChart(
                    "Ops",
                    new ChartInput.Pie(
                        "ShareChart",
                        anchor,
                        new ChartInput.Title.Formula("C1"),
                        null,
                        null,
                        null,
                        null,
                        120,
                        List.of(secondSeries)))));

    assertEquals("Ops", lineCommand.sheetName());
    ExcelChartDefinition.Line lineChart =
        assertInstanceOf(ExcelChartDefinition.Line.class, lineCommand.chart());
    assertEquals(new ExcelChartDefinition.Title.Text("Trend"), lineChart.title());
    assertEquals(new ExcelChartDefinition.Legend.Hidden(), lineChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, lineChart.displayBlanksAs());
    assertFalse(lineChart.plotOnlyVisibleCells());
    assertTrue(lineChart.varyColors());
    assertEquals(
        new ExcelChartDefinition.Series(
            new ExcelChartDefinition.Title.Formula("B1"),
            new ExcelChartDefinition.DataSource("A2:A4"),
            new ExcelChartDefinition.DataSource("B2:B4")),
        lineChart.series().getFirst());

    ExcelChartDefinition.Pie pieChart =
        assertInstanceOf(ExcelChartDefinition.Pie.class, pieCommand.chart());
    assertEquals(new ExcelChartDefinition.Title.Formula("C1"), pieChart.title());
    assertEquals(
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT), pieChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, pieChart.displayBlanksAs());
    assertTrue(pieChart.plotOnlyVisibleCells());
    assertFalse(pieChart.varyColors());
    assertEquals(120, pieChart.firstSliceAngle());
    assertEquals(
        new ExcelChartDefinition.Series(
            new ExcelChartDefinition.Title.Text("Actual"),
            new ExcelChartDefinition.DataSource("ChartCategories"),
            new ExcelChartDefinition.DataSource("ChartActual")),
        pieChart.series().getFirst());
  }

  @Test
  void directChartHelperConversionsCoverStandaloneSwitchBranches() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(2, 3, 0, 0),
            new DrawingMarkerInput(9, 16, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    ChartInput.Series barSeries =
        new ChartInput.Series(
            new ChartInput.Title.None(),
            new ChartInput.DataSource("Summary!$A$2:$A$4"),
            new ChartInput.DataSource("Summary!$B$2:$B$4"));
    ChartInput barInput =
        new ChartInput.Bar(
            "OpsBar",
            anchor,
            new ChartInput.Title.Formula("Summary!$B$1"),
            new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            true,
            ExcelChartBarDirection.BAR,
            List.of(barSeries));
    ChartInput lineInput =
        new ChartInput.Line(
            "OpsLine",
            anchor,
            new ChartInput.Title.Text("Trend"),
            new ChartInput.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            true,
            false,
            List.of(
                new ChartInput.Series(
                    new ChartInput.Title.Formula("Summary!$C$1"),
                    new ChartInput.DataSource("ChartCategories"),
                    new ChartInput.DataSource("ChartActual"))));
    ChartInput pieInput =
        new ChartInput.Pie(
            "OpsPie",
            anchor,
            new ChartInput.Title.Text("Share"),
            new ChartInput.Legend.Visible(ExcelChartLegendPosition.LEFT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            true,
            90,
            List.of(
                new ChartInput.Series(
                    new ChartInput.Title.Text("Actual"),
                    new ChartInput.DataSource("ChartCategories"),
                    new ChartInput.DataSource("ChartActual"))));

    ExcelChartDefinition.Bar bar =
        assertInstanceOf(
            ExcelChartDefinition.Bar.class,
            WorkbookCommandConverter.toExcelChartDefinition(barInput));
    ExcelChartDefinition.Line line =
        assertInstanceOf(
            ExcelChartDefinition.Line.class,
            WorkbookCommandConverter.toExcelChartDefinition(lineInput));
    ExcelChartDefinition.Pie pie =
        assertInstanceOf(
            ExcelChartDefinition.Pie.class,
            WorkbookCommandConverter.toExcelChartDefinition(pieInput));
    WorkbookOperation barOperation = new WorkbookOperation.SetChart("Ops", barInput);
    WorkbookCommand.SetChart barCommand =
        assertInstanceOf(
            WorkbookCommand.SetChart.class, WorkbookCommandConverter.toCommand(barOperation));

    assertEquals(new ExcelChartDefinition.Title.Formula("Summary!$B$1"), bar.title());
    assertEquals(
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP), bar.legend());
    assertEquals(ExcelChartDisplayBlanksAs.SPAN, bar.displayBlanksAs());
    assertFalse(bar.plotOnlyVisibleCells());
    assertTrue(bar.varyColors());
    assertEquals(ExcelChartBarDirection.BAR, bar.barDirection());
    assertEquals(
        new ExcelChartDefinition.Series(
            new ExcelChartDefinition.Title.None(),
            new ExcelChartDefinition.DataSource("Summary!$A$2:$A$4"),
            new ExcelChartDefinition.DataSource("Summary!$B$2:$B$4")),
        bar.series().getFirst());
    assertEquals(new ExcelChartDefinition.Title.Text("Trend"), line.title());
    assertEquals(new ExcelChartDefinition.Legend.Hidden(), line.legend());
    assertEquals(
        new ExcelChartDefinition.Title.Formula("Summary!$C$1"), line.series().getFirst().title());
    assertEquals(new ExcelChartDefinition.Title.Text("Share"), pie.title());
    assertEquals(
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.LEFT), pie.legend());
    assertTrue(pie.varyColors());
    assertEquals(90, pie.firstSliceAngle());
    assertEquals("Ops", barCommand.sheetName());
    assertEquals(bar, barCommand.chart());
  }

  @Test
  void convertsAdvancedStyleAndConditionalFormattingPayloads() {
    ExcelCellStyle style =
        WorkbookCommandConverter.toExcelCellStyle(
            new CellStyleInput(
                null,
                null,
                new CellFontInput(null, null, null, null, null, 2, null, 0.4d, true, null),
                new CellFillInput(
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    new CellGradientFillInput(
                        " path ",
                        null,
                        0.1d,
                        0.2d,
                        0.3d,
                        0.4d,
                        List.of(
                            new CellGradientStopInput(0.0d, new ColorInput("#112233")),
                            new CellGradientStopInput(1.0d, new ColorInput(null, 5, null, 0.2d))))),
                new CellBorderInput(
                    null, new CellBorderSideInput(null, null, 1, null, 0.15d), null, null, null),
                null));

    assertEquals(
        new ExcelCellFont(null, null, null, null, new ExcelColor(null, 2, null, 0.4d), true, null),
        style.font());
    assertEquals(
        new ExcelCellFill(
            null,
            null,
            null,
            new ExcelGradientFill(
                "PATH",
                null,
                0.1d,
                0.2d,
                0.3d,
                0.4d,
                List.of(
                    new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                    new ExcelGradientStop(1.0d, new ExcelColor(null, 5, null, 0.2d))))),
        style.fill());
    assertEquals(
        new ExcelBorder(
            null,
            new ExcelBorderSide(null, new ExcelColor(null, 1, null, 0.15d)),
            null,
            null,
            null),
        style.border());

    assertEquals(
        new ExcelConditionalFormattingRule.ColorScaleRule(
            List.of(
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.MIN, null, null),
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.MAX, null, null)),
            List.of(new ExcelColor("#112233"), new ExcelColor("#AABBCC")),
            true),
        WorkbookCommandConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.ColorScaleRule(
                true,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                List.of(new ColorInput("#112233"), new ColorInput("#AABBCC")))));

    assertEquals(
        new ExcelConditionalFormattingRule.DataBarRule(
            new ExcelColor(null, 4, null, 0.25d),
            true,
            10,
            90,
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.PERCENTILE, null, 90.0d),
            false),
        WorkbookCommandConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                new ColorInput(null, 4, null, 0.25d),
                true,
                10,
                90,
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.PERCENTILE, null, 90.0d))));

    assertEquals(
        new ExcelConditionalFormattingRule.IconSetRule(
            ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
            false,
            true,
            List.of(
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.PERCENT, null, 33.0d),
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.PERCENT, null, 67.0d)),
            true),
        WorkbookCommandConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.IconSetRule(
                true,
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                true,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 33.0d),
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 67.0d)))));

    assertEquals(
        new ExcelConditionalFormattingRule.Top10Rule(7, true, false, false, null),
        WorkbookCommandConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.Top10Rule(false, 7, true, false, null)));
  }

  @Test
  void executorContextHelpersReturnNoFalseMetadataForWorkbookProtectionOperations() {
    WorkbookOperation.SetWorkbookProtection setProtection =
        new WorkbookOperation.SetWorkbookProtection(
            new WorkbookProtectionInput(true, false, true, "book-secret", null));
    WorkbookOperation.ClearWorkbookProtection clearProtection =
        new WorkbookOperation.ClearWorkbookProtection();

    assertEquals("SET_WORKBOOK_PROTECTION", setProtection.operationType());
    assertEquals("CLEAR_WORKBOOK_PROTECTION", clearProtection.operationType());
    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(setProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.sheetNameFor(setProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.addressFor(setProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.rangeFor(setProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            setProtection, new IllegalStateException()));

    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(clearProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.sheetNameFor(clearProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.addressFor(clearProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.rangeFor(clearProtection, new IllegalStateException()));
    assertNull(
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            clearProtection, new IllegalStateException()));
  }

  private static WorkbookOperation.SetAutofilter advancedAutofilterOperation() {
    return new WorkbookOperation.SetAutofilter(
        "Budget",
        "A1:F9",
        List.of(
            new AutofilterFilterColumnInput(
                0L,
                null,
                new AutofilterFilterCriterionInput.Values(List.of("Queued", "Ready"), true)),
            new AutofilterFilterColumnInput(
                1L,
                false,
                new AutofilterFilterCriterionInput.Custom(
                    true,
                    List.of(
                        new AutofilterFilterCriterionInput.CustomConditionInput(
                            "greaterThan", "5")))),
            new AutofilterFilterColumnInput(
                2L, true, new AutofilterFilterCriterionInput.Dynamic("TODAY", 1.0d, 2.0d)),
            new AutofilterFilterColumnInput(
                3L, true, new AutofilterFilterCriterionInput.Top10(10, true, false)),
            new AutofilterFilterColumnInput(
                4L,
                true,
                new AutofilterFilterCriterionInput.Color(
                    false, new ColorInput(null, 3, null, 0.25d))),
            new AutofilterFilterColumnInput(
                5L, true, new AutofilterFilterCriterionInput.Icon("3TrafficLights1", 2))),
        new AutofilterSortStateInput(
            "A2:F9",
            null,
            true,
            null,
            List.of(
                new AutofilterSortConditionInput(
                    "B2:B9", true, null, new ColorInput("#AABBCC"), null),
                new AutofilterSortConditionInput("C2:C9", false, "ICON", null, 2))));
  }

  private static List<ExcelAutofilterFilterColumn> expectedAutofilterCriteria() {
    return List.of(
        new ExcelAutofilterFilterColumn(
            0L, true, new ExcelAutofilterFilterCriterion.Values(List.of("Queued", "Ready"), true)),
        new ExcelAutofilterFilterColumn(
            1L,
            false,
            new ExcelAutofilterFilterCriterion.Custom(
                true,
                List.of(new ExcelAutofilterFilterCriterion.CustomCondition("greaterThan", "5")))),
        new ExcelAutofilterFilterColumn(
            2L, true, new ExcelAutofilterFilterCriterion.Dynamic("TODAY", 1.0d, 2.0d)),
        new ExcelAutofilterFilterColumn(
            3L, true, new ExcelAutofilterFilterCriterion.Top10(10, true, false)),
        new ExcelAutofilterFilterColumn(
            4L,
            true,
            new ExcelAutofilterFilterCriterion.Color(false, new ExcelColor(null, 3, null, 0.25d))),
        new ExcelAutofilterFilterColumn(
            5L, true, new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", 2)));
  }

  private static ExcelAutofilterSortState expectedAutofilterSortState() {
    return new ExcelAutofilterSortState(
        "A2:F9",
        false,
        true,
        "",
        List.of(
            new ExcelAutofilterSortCondition("B2:B9", true, "", new ExcelColor("#AABBCC"), null),
            new ExcelAutofilterSortCondition("C2:C9", false, "ICON", null, 2)));
  }
}
