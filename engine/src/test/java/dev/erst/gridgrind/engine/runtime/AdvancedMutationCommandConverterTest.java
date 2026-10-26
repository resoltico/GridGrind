package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterion;
import dev.erst.gridgrind.excel.ExcelAutofilterSortCondition;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableDefinition;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused command-converter coverage for advanced workbook-core mutation payloads. */
class AdvancedMutationCommandConverterTest {
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC";

  @Test
  void convertsAdvancedCommentProtectionAutofilterAndNamedRangeOperations() {
    WorkbookAnnotationCommand.SetComment commentCommand =
        assertInstanceOf(
            WorkbookAnnotationCommand.SetComment.class,
            WorkbookCommandConverter.toCommand(
                new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B4")),
                new CellMutationAction.SetComment(
                    new CommentSpec(
                        text("Ada Lovelace"),
                        "GridGrind",
                        true,
                        java.util.Optional.of(
                            List.of(
                                new RichTextRunSpec(text("Ada"), Optional.empty()),
                                new RichTextRunSpec(text(" Lovelace"), Optional.empty()))),
                        java.util.Optional.of(new CommentAnchorSpec(1, 2, 4, 6))))));
    assertEquals(
        new ResolvedComment(
            "Ada Lovelace",
            "GridGrind",
            true,
            Optional.of(
                new ResolvedRichText(
                    List.of(
                        new ResolvedRichTextRun("Ada", Optional.empty()),
                        new ResolvedRichTextRun(" Lovelace", Optional.empty())))),
            Optional.of(new CommentAnchorSpec(1, 2, 4, 6))),
        commentCommand.comment());

    WorkbookSheetCommand.SetWorkbookProtection protectionCommand =
        assertInstanceOf(
            WorkbookSheetCommand.SetWorkbookProtection.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookSelector.Current(),
                new WorkbookMutationAction.SetWorkbookProtection(
                    new WorkbookProtectionSpec(
                        false, true, false, Optional.empty(), Optional.empty()))));
    assertEquals(
        new ExcelWorkbookProtectionSettings(false, true, false, Optional.empty(), Optional.empty()),
        protectionCommand.protection());

    assertInstanceOf(
        WorkbookSheetCommand.ClearWorkbookProtection.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookSelector.Current(), new WorkbookMutationAction.ClearWorkbookProtection()));

    WorkbookTabularCommand.SetAutofilter simpleAutofilter =
        assertInstanceOf(
            WorkbookTabularCommand.SetAutofilter.class,
            WorkbookCommandConverter.toCommand(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:C4")),
                new StructuredMutationAction.SetAutofilter()));
    assertEquals(List.of(), simpleAutofilter.criteria());
    assertEquals(Optional.empty(), simpleAutofilter.sortState());
    assertEquals(List.of(), new StructuredMutationAction.SetAutofilter().criteria());

    WorkbookTabularCommand.SetAutofilter advancedAutofilter =
        assertInstanceOf(
            WorkbookTabularCommand.SetAutofilter.class,
            WorkbookCommandConverter.toCommand(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:F9")),
                advancedAutofilterAction()));
    assertEquals(expectedAutofilterCriteria(), advancedAutofilter.criteria());
    assertEquals(Optional.of(expectedAutofilterSortState()), advancedAutofilter.sortState());

    WorkbookMetadataCommand.SetNamedRange namedRangeCommand =
        assertInstanceOf(
            WorkbookMetadataCommand.SetNamedRange.class,
            WorkbookCommandConverter.toCommand(
                new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetExpr")),
                new StructuredMutationAction.SetNamedRange(
                    new DefinedName("BudgetExpr"),
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.formula("SUM(Budget!A1:A3)"))));
    assertEquals(
        new ExcelNamedRangeDefinition(
            "BudgetExpr",
            new ExcelNamedRangeScope.WorkbookScope(),
            ExcelNamedRangeTarget.formula("SUM(Budget!A1:A3)")),
        namedRangeCommand.definition());

    WorkbookMutationAction.SetSheetProtection sheetProtection =
        new WorkbookMutationAction.SetSheetProtection(
            new SheetProtectionSettings(
                true, false, false, false, false, false, false, false, false, false, false, false,
                false, false, false),
            Optional.of(new dev.erst.gridgrind.core.source.SecretReference("sheet-password")));
    assertEquals(
        Optional.of(new dev.erst.gridgrind.core.source.SecretReference("sheet-password")),
        sheetProtection.passwordRef());
    org.junit.jupiter.api.Assertions.assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookMutationAction.SetSheetProtection(
                new SheetProtectionSettings(
                    true, false, false, false, false, false, false, false, false, false, false,
                    false, false, false, false),
                Optional.of(new dev.erst.gridgrind.core.source.SecretReference(" "))));
  }

  @Test
  void convertsDrawingMutationOperations() {
    DrawingAnchorSpec.TwoCell anchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(1, 2, 3, 4),
            new DrawingMarkerSpec(4, 6, 7, 8),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    PictureDataSpec pictureData =
        new PictureDataSpec(ExcelPictureFormat.PNG, binary(PNG_PIXEL_BASE64));

    WorkbookDrawingCommand.SetPicture pictureCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetPicture.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetPicture(
                    new PictureSpec(
                        "OpsPicture", pictureData, anchor, Optional.of(text("Queue preview"))))));
    WorkbookDrawingCommand.SetSignatureLine signatureLineCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetSignatureLine.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetSignatureLine(
                    new SignatureLineSpec(
                        "OpsSignature",
                        anchor,
                        false,
                        java.util.Optional.of("Review before signing."),
                        java.util.Optional.of("Ada Lovelace"),
                        java.util.Optional.of("Finance"),
                        java.util.Optional.of("ada@example.com"),
                        java.util.Optional.empty(),
                        java.util.Optional.of("invalid"),
                        java.util.Optional.of(pictureData)))));
    WorkbookDrawingCommand.SetShape shapeCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetShape.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetShape(
                    new ShapeSpec.SimpleShape(
                        "OpsShape", anchor, "rect", Optional.of(text("Queue"))))));
    WorkbookDrawingCommand.SetEmbeddedObject embeddedObjectCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetEmbeddedObject.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetEmbeddedObject(
                    new EmbeddedObjectSpec(
                        "OpsEmbed",
                        "Payload",
                        "payload.txt",
                        "payload.txt",
                        binary("cGF5bG9hZA=="),
                        pictureData,
                        anchor))));
    WorkbookDrawingCommand.SetDrawingObjectAnchor moveCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetDrawingObjectAnchor.class,
            WorkbookCommandConverter.toCommand(
                new DrawingObjectSelector.ByName(new SheetName("Ops"), "OpsPicture"),
                new DrawingMutationAction.SetDrawingObjectAnchor(anchor)));
    WorkbookDrawingCommand.DeleteDrawingObject deleteCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.DeleteDrawingObject.class,
            WorkbookCommandConverter.toCommand(
                new DrawingObjectSelector.ByName(new SheetName("Ops"), "OpsPicture"),
                new DrawingMutationAction.DeleteDrawingObject()));

    assertEquals(
        new ExcelPictureDefinition(
            "OpsPicture",
            new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
            ExcelPictureFormat.PNG,
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 3, 4),
                new DrawingMarkerSpec(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
            Optional.of("Queue preview")),
        pictureCommand.picture());
    assertEquals(
        new ExcelSignatureLineDefinition(
            "OpsSignature",
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 3, 4),
                new DrawingMarkerSpec(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
            false,
            "Review before signing.",
            "Ada Lovelace",
            "Finance",
            "ada@example.com",
            null,
            "invalid",
            Optional.of(ExcelPictureFormat.PNG),
            Optional.of(
                new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)))),
        signatureLineCommand.signatureLine());
    assertEquals(
        new ExcelShapeDefinition.SimpleShape(
            "OpsShape",
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 3, 4),
                new DrawingMarkerSpec(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
            "rect",
            Optional.of("Queue")),
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
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 3, 4),
                new DrawingMarkerSpec(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE)),
        embeddedObjectCommand.embeddedObject());
    assertEquals("OpsPicture", moveCommand.objectName());
    assertEquals("OpsPicture", deleteCommand.objectName());
  }

  @Test
  void convertsArrayFormulaMutations() {
    WorkbookCellCommand.SetArrayFormula setArrayFormula =
        assertInstanceOf(
            WorkbookCellCommand.SetArrayFormula.class,
            WorkbookCommandConverter.toCommand(
                new RangeSelector.ByRange(new SheetName("Calc"), new RangeRef("D2:D4")),
                new CellMutationAction.SetArrayFormula(
                    new ArrayFormulaSpec(text("{=B2:B4*C2:C4}")))));
    WorkbookCellCommand.ClearArrayFormula clearArrayFormula =
        assertInstanceOf(
            WorkbookCellCommand.ClearArrayFormula.class,
            WorkbookCommandConverter.toCommand(
                new CellSelector.ByAddress(new SheetName("Calc"), new CellAddress("D3")),
                new CellMutationAction.ClearArrayFormula()));

    assertEquals("Calc", setArrayFormula.sheetName());
    assertEquals("D2:D4", setArrayFormula.range());
    assertEquals("B2:B4*C2:C4", setArrayFormula.formula().formula());
    assertEquals("Calc", clearArrayFormula.sheetName());
    assertEquals("D3", clearArrayFormula.address());
  }

  @Test
  void convertsCustomXmlImportMutations() {
    WorkbookMetadataCommand.ImportCustomXmlMapping command =
        assertInstanceOf(
            WorkbookMetadataCommand.ImportCustomXmlMapping.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookSelector.Current(),
                new StructuredMutationAction.ImportCustomXmlMapping(
                    new CustomXmlImportSpec(
                        new CustomXmlMappingLocator(
                            java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                        text("<CORSO><NOME>Grid</NOME></CORSO>")))));

    assertEquals(
        new ExcelCustomXmlImportDefinition(
            new CustomXmlMappingLocator(
                java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
            "<CORSO><NOME>Grid</NOME></CORSO>"),
        command.mapping());
  }

  @Test
  void convertsPivotTableMutationOperations() {
    WorkbookTabularCommand.SetPivotTable setPivotTable =
        assertInstanceOf(
            WorkbookTabularCommand.SetPivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", new SheetName("Report")),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableSpec(
                        "Sales Pivot 2026",
                        new SheetName("Report"),
                        new PivotTableSpec.PivotSourceSpec.Range(
                            new SheetName("Data"), new RangeRef("A1:D5")),
                        new PivotTableSpec.Anchor(new CellAddress("C5")),
                        List.of("Region"),
                        List.of("Stage"),
                        List.of("Owner"),
                        List.of(
                            new PivotTableSpec.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Amount",
                                Optional.of("#,##0.00")))))));
    WorkbookTabularCommand.SetPivotTable setPivotTableFromNamedRange =
        assertInstanceOf(
            WorkbookTabularCommand.SetPivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Named Source Pivot", new SheetName("Report")),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableSpec(
                        "Named Source Pivot",
                        new SheetName("Report"),
                        new PivotTableSpec.PivotSourceSpec.NamedRange(
                            new DefinedName("PivotSource")),
                        new PivotTableSpec.Anchor(new CellAddress("A3")),
                        List.of("Region"),
                        List.of(),
                        List.of(),
                        List.of(
                            new PivotTableSpec.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                Optional.empty()))))));
    WorkbookTabularCommand.SetPivotTable setPivotTableFromTable =
        assertInstanceOf(
            WorkbookTabularCommand.SetPivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Table Source Pivot", new SheetName("Report")),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableSpec(
                        "Table Source Pivot",
                        new SheetName("Report"),
                        new PivotTableSpec.PivotSourceSpec.Table(new TableName("SalesTable2026")),
                        new PivotTableSpec.Anchor(new CellAddress("G4")),
                        List.of("Region"),
                        List.of(),
                        List.of(),
                        List.of(
                            new PivotTableSpec.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                Optional.empty()))))));
    WorkbookTabularCommand.DeletePivotTable deletePivotTable =
        assertInstanceOf(
            WorkbookTabularCommand.DeletePivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", new SheetName("Report")),
                new StructuredMutationAction.DeletePivotTable()));

    assertEquals(
        new ExcelPivotTableDefinition(
            "Sales Pivot 2026",
            "Report",
            new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
            new ExcelPivotTableDefinition.Anchor("C5"),
            List.of("Region"),
            List.of("Stage"),
            List.of("Owner"),
            List.of(
                new ExcelPivotTableDefinition.DataField(
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Amount",
                    Optional.of("#,##0.00")))),
        setPivotTable.definition());
    assertEquals(
        new ExcelPivotTableDefinition.Source.NamedRange("PivotSource"),
        setPivotTableFromNamedRange.definition().source());
    assertEquals(
        new ExcelPivotTableDefinition.Source.Table("SalesTable2026"),
        setPivotTableFromTable.definition().source());
    assertEquals("Sales Pivot 2026", deletePivotTable.name());
    assertEquals("Report", deletePivotTable.sheetName());
  }

  @Test
  void convertsChartMutationOperationsAcrossSupportedFamilies() {
    DrawingAnchorSpec.TwoCell anchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(1, 2, 3, 4),
            new DrawingMarkerSpec(7, 12, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ChartSeriesSpec firstSeries = chartSeries(new ChartTitleSpec.Formula("B1"), "A2:A4", "B2:B4");
    ChartSeriesSpec secondSeries =
        chartSeries(new ChartTitleSpec.Text(text("Actual")), "ChartCategories", "ChartActual");

    WorkbookDrawingCommand.SetChart lineCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetChart(
                    chartInput(
                        "TrendChart",
                        anchor,
                        new ChartTitleSpec.Text(text("Trend")),
                        new ChartLegendSpec.Hidden(),
                        ExcelChartDisplayBlanksAs.ZERO,
                        false,
                        new ChartPlotSpec.Line(
                            true, ExcelChartGrouping.STANDARD, List.of(firstSeries))))));
    WorkbookDrawingCommand.SetChart pieCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetChart(
                    chartInput(
                        "ShareChart",
                        anchor,
                        new ChartTitleSpec.Formula("C1"),
                        null,
                        null,
                        null,
                        new ChartPlotSpec.Pie(false, Optional.of(120), List.of(secondSeries))))));

    assertEquals("Ops", lineCommand.sheetName());
    ResolvedChart lineChart = lineCommand.chart();
    ResolvedChart.Line linePlot =
        assertInstanceOf(ResolvedChart.Line.class, lineChart.plots().getFirst());
    assertEquals(new ResolvedChartTitle.Text("Trend"), lineChart.title());
    assertEquals(new ChartLegendSpec.Hidden(), lineChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, lineChart.displayBlanksAs());
    assertFalse(lineChart.plotOnlyVisibleCells());
    assertTrue(linePlot.varyColors());
    assertEquals(
        chartDefinitionSeries(new ChartTitleSpec.Formula("B1"), "A2:A4", "B2:B4"),
        linePlot.series().getFirst());

    ResolvedChart pieChart = pieCommand.chart();
    ResolvedChart.Pie piePlot =
        assertInstanceOf(ResolvedChart.Pie.class, pieChart.plots().getFirst());
    assertEquals(new ChartTitleSpec.Formula("C1"), pieChart.title());
    assertEquals(new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT), pieChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, pieChart.displayBlanksAs());
    assertTrue(pieChart.plotOnlyVisibleCells());
    assertFalse(piePlot.varyColors());
    assertEquals(Optional.of(120), piePlot.firstSliceAngle());
    assertEquals(
        chartDefinitionSeries(
            new ResolvedChartTitle.Text("Actual"), "ChartCategories", "ChartActual"),
        piePlot.series().getFirst());
  }

  @Test
  void directChartHelperConversionsCoverStandaloneSwitchBranches() {
    DrawingAnchorSpec.TwoCell anchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(2, 3, 0, 0),
            new DrawingMarkerSpec(9, 16, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    ChartSeriesSpec barSeries =
        chartSeries(new ChartTitleSpec.None(), "Summary!$A$2:$A$4", "Summary!$B$2:$B$4");
    ChartSpec barInput =
        chartInput(
            "OpsBar",
            anchor,
            new ChartTitleSpec.Formula("Summary!$B$1"),
            new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            new ChartPlotSpec.Bar(
                true,
                ExcelChartBarDirection.BAR,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                List.of(barSeries)));
    ChartSpec lineInput =
        chartInput(
            "OpsLine",
            anchor,
            new ChartTitleSpec.Text(text("Trend")),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            true,
            new ChartPlotSpec.Line(
                false,
                ExcelChartGrouping.STANDARD,
                List.of(
                    chartSeries(
                        new ChartTitleSpec.Formula("Summary!$C$1"),
                        "ChartCategories",
                        "ChartActual"))));
    ChartSpec pieInput =
        chartInput(
            "OpsPie",
            anchor,
            new ChartTitleSpec.Text(text("Share")),
            new ChartLegendSpec.Visible(ExcelChartLegendPosition.LEFT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            new ChartPlotSpec.Pie(
                true,
                Optional.of(90),
                List.of(
                    chartSeries(
                        new ChartTitleSpec.Text(text("Actual")),
                        "ChartCategories",
                        "ChartActual"))));

    ResolvedChart bar = WorkbookCommandDrawingInputConverter.toResolvedChart(barInput);
    ResolvedChart.Line linePlot =
        assertInstanceOf(
            ResolvedChart.Line.class,
            WorkbookCommandDrawingInputConverter.toResolvedChart(lineInput).plots().getFirst());
    ResolvedChart.Pie piePlot =
        assertInstanceOf(
            ResolvedChart.Pie.class,
            WorkbookCommandDrawingInputConverter.toResolvedChart(pieInput).plots().getFirst());
    WorkbookDrawingCommand.SetChart barCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetChart(barInput)));

    assertEquals(new ChartTitleSpec.Formula("Summary!$B$1"), bar.title());
    assertEquals(new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP), bar.legend());
    assertEquals(ExcelChartDisplayBlanksAs.SPAN, bar.displayBlanksAs());
    assertFalse(bar.plotOnlyVisibleCells());
    ResolvedChart.Bar barPlot = assertInstanceOf(ResolvedChart.Bar.class, bar.plots().getFirst());
    assertTrue(barPlot.varyColors());
    assertEquals(ExcelChartBarDirection.BAR, barPlot.barDirection());
    assertEquals(ExcelChartBarGrouping.CLUSTERED, barPlot.grouping());
    assertEquals(
        chartDefinitionSeries(new ChartTitleSpec.None(), "Summary!$A$2:$A$4", "Summary!$B$2:$B$4"),
        barPlot.series().getFirst());
    ResolvedChart line = WorkbookCommandDrawingInputConverter.toResolvedChart(lineInput);
    assertEquals(new ResolvedChartTitle.Text("Trend"), line.title());
    assertEquals(new ChartLegendSpec.Hidden(), line.legend());
    assertEquals(new ChartTitleSpec.Formula("Summary!$C$1"), linePlot.series().getFirst().title());
    ResolvedChart pie = WorkbookCommandDrawingInputConverter.toResolvedChart(pieInput);
    assertEquals(new ResolvedChartTitle.Text("Share"), pie.title());
    assertEquals(new ChartLegendSpec.Visible(ExcelChartLegendPosition.LEFT), pie.legend());
    assertTrue(piePlot.varyColors());
    assertEquals(Optional.of(90), piePlot.firstSliceAngle());
    assertEquals("Ops", barCommand.sheetName());
    assertEquals(bar, barCommand.chart());
  }

  @Test
  void convertsAdvancedConditionalFormattingPayloads() {
    assertEquals(
        new ExcelConditionalFormattingRule.ColorScaleRule(
            List.of(
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.MIN, null, null),
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.MAX, null, null)),
            List.of(ColorSpec.rgb(new RgbHex("#112233")), ColorSpec.rgb(new RgbHex("#AABBCC"))),
            true),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleSpec.ColorScaleRule(
                true,
                List.of(
                    new ConditionalFormattingThresholdSpec.Min(),
                    new ConditionalFormattingThresholdSpec.Max()),
                List.of(
                    ColorSpec.rgb(new RgbHex("#112233")), ColorSpec.rgb(new RgbHex("#AABBCC"))))));

    assertEquals(
        new ExcelConditionalFormattingRule.DataBarRule(
            ColorSpec.theme(4, 0.25d),
            true,
            10,
            90,
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.PERCENTILE, null, 90.0d),
            false),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleSpec.DataBarRule(
                false,
                ColorSpec.theme(4, 0.25d),
                true,
                10,
                90,
                new ConditionalFormattingThresholdSpec.Numeric(0.0d),
                new ConditionalFormattingThresholdSpec.Percentile(90.0d))));

    ExcelConditionalFormattingRule.ColorScaleRule formulaThresholdRule =
        assertInstanceOf(
            ExcelConditionalFormattingRule.ColorScaleRule.class,
            WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
                new ConditionalFormattingRuleSpec.ColorScaleRule(
                    false,
                    List.of(
                        new ConditionalFormattingThresholdSpec.Formula("A1"),
                        new ConditionalFormattingThresholdSpec.Percent(100.0d)),
                    List.of(
                        ColorSpec.rgb(new RgbHex("#112233")),
                        ColorSpec.rgb(new RgbHex("#AABBCC"))))));
    assertEquals(
        new ExcelConditionalFormattingThreshold(
            ExcelConditionalFormattingThresholdType.FORMULA, "A1", null),
        formulaThresholdRule.thresholds().getFirst());

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
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleSpec.IconSetRule(
                true,
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                true,
                List.of(
                    new ConditionalFormattingThresholdSpec.Percent(0.0d),
                    new ConditionalFormattingThresholdSpec.Percent(33.0d),
                    new ConditionalFormattingThresholdSpec.Percent(67.0d)))));

    assertEquals(
        new ExcelConditionalFormattingRule.Top10Rule(7, true, false, true, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleSpec.Top10Rule(true, 7, true, false, Optional.empty())));
  }

  @Test
  void executorContextHelpersReturnNoFalseMetadataForWorkbookProtectionOperations() {
    WorkbookMutationAction.SetWorkbookProtection setProtection =
        new WorkbookMutationAction.SetWorkbookProtection(
            new WorkbookProtectionSpec(
                true,
                false,
                true,
                Optional.of(new dev.erst.gridgrind.core.source.SecretReference("book-password")),
                Optional.empty()));
    WorkbookMutationAction.ClearWorkbookProtection clearProtection =
        new WorkbookMutationAction.ClearWorkbookProtection();

    assertEquals("SET_WORKBOOK_PROTECTION", setProtection.actionType());
    assertEquals("CLEAR_WORKBOOK_PROTECTION", clearProtection.actionType());
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.formulaFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.sheetNameFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.addressFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.rangeFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));

    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.formulaFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.sheetNameFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.addressFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.rangeFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
  }

  private static StructuredMutationAction.SetAutofilter advancedAutofilterAction() {
    return new StructuredMutationAction.SetAutofilter(
        List.of(
            AutofilterFilterColumnSpec.visibleButton(
                0L, new AutofilterFilterCriterionSpec.Values(List.of("Queued", "Ready"), true)),
            new AutofilterFilterColumnSpec(
                1L,
                false,
                new AutofilterFilterCriterionSpec.Custom(
                    true,
                    List.of(
                        new AutofilterFilterCriterionSpec.CustomConditionSpec(
                            "greaterThan", "5")))),
            new AutofilterFilterColumnSpec(
                2L,
                true,
                new AutofilterFilterCriterionSpec.Dynamic(
                    "TODAY", java.util.Optional.of(1.0d), java.util.Optional.of(2.0d))),
            new AutofilterFilterColumnSpec(
                3L, true, new AutofilterFilterCriterionSpec.Top10(10, true, false)),
            new AutofilterFilterColumnSpec(
                4L,
                true,
                new AutofilterFilterCriterionSpec.Color(false, ColorSpec.theme(3, 0.25d))),
            new AutofilterFilterColumnSpec(
                5L, true, new AutofilterFilterCriterionSpec.Icon("3TrafficLights1", 2))),
        AutofilterSortStateSpec.withoutSortMethod(
            new RangeRef("A2:F9"),
            false,
            true,
            List.of(
                new AutofilterSortConditionSpec.CellColor(
                    new RangeRef("B2:B9"), true, ColorSpec.rgb(new RgbHex("#AABBCC"))),
                new AutofilterSortConditionSpec.FontColor(
                    new RangeRef("C2:C9"), false, ColorSpec.theme(4, 0.15d)),
                new AutofilterSortConditionSpec.Icon(new RangeRef("C2:C9"), false, 2))));
  }

  private static TextSourceSpec text(String value) {
    return TextSourceSpec.inline(value);
  }

  private static BinarySourceSpec binary(String value) {
    return BinarySourceSpec.inlineBase64(value);
  }

  private static ChartSpec chartInput(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      ChartTitleSpec title,
      ChartLegendSpec legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      ChartPlotSpec plot) {
    return new ChartSpec(
        name,
        anchor,
        title == null ? new ChartTitleSpec.None() : title,
        legend == null ? new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT) : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        List.of(plot));
  }

  private static ChartSeriesSpec chartSeries(
      ChartTitleSpec title, String categoriesFormula, String valuesFormula) {
    return new ChartSeriesSpec(
        title,
        new ChartDataSourceSpec.Reference(categoriesFormula),
        new ChartDataSourceSpec.Reference(valuesFormula),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static ResolvedChart.Series chartDefinitionSeries(
      ResolvedChartTitle title, String categoriesFormula, String valuesFormula) {
    return new ResolvedChart.Series(
        title,
        new ChartDataSourceSpec.Reference(categoriesFormula),
        new ChartDataSourceSpec.Reference(valuesFormula),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
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
            4L, true, new ExcelAutofilterFilterCriterion.Color(false, ColorSpec.theme(3, 0.25d))),
        new ExcelAutofilterFilterColumn(
            5L, true, new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", 2)));
  }

  private static ExcelAutofilterSortState expectedAutofilterSortState() {
    return new ExcelAutofilterSortState(
        "A2:F9",
        false,
        true,
        java.util.Optional.empty(),
        List.of(
            new ExcelAutofilterSortCondition.CellColor(
                "B2:B9", true, ColorSpec.rgb(new RgbHex("#AABBCC"))),
            new ExcelAutofilterSortCondition.FontColor("C2:C9", false, ColorSpec.theme(4, 0.15d)),
            new ExcelAutofilterSortCondition.Icon("C2:C9", false, 2)));
  }
}
