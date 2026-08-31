package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.BorderSideInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellGradientFillInput;
import dev.erst.gridgrind.contract.dto.CellGradientStopInput;
import dev.erst.gridgrind.contract.dto.CellStylePatchInput;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CommentAnchorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.*;
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
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentAnchor;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
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
                new CellSelector.ByAddress("Budget", "B4"),
                new CellMutationAction.SetComment(
                    new CommentInput(
                        text("Ada Lovelace"),
                        "GridGrind",
                        true,
                        java.util.Optional.of(
                            List.of(
                                new RichTextRunInput(text("Ada"), Optional.empty()),
                                new RichTextRunInput(text(" Lovelace"), Optional.empty()))),
                        java.util.Optional.of(new CommentAnchorInput(1, 2, 4, 6))))));
    assertEquals(
        new ExcelComment(
            "Ada Lovelace",
            "GridGrind",
            true,
            Optional.of(
                new ExcelRichText(
                    List.of(
                        new ExcelRichTextRun("Ada", Optional.empty()),
                        new ExcelRichTextRun(" Lovelace", Optional.empty())))),
            Optional.of(new ExcelCommentAnchor(1, 2, 4, 6))),
        commentCommand.comment());

    WorkbookSheetCommand.SetWorkbookProtection protectionCommand =
        assertInstanceOf(
            WorkbookSheetCommand.SetWorkbookProtection.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookSelector.Current(),
                new WorkbookMutationAction.SetWorkbookProtection(
                    new WorkbookProtectionInput(
                        false,
                        true,
                        false,
                        Optional.of("book-secret"),
                        Optional.of("review-secret")))));
    assertEquals(
        new ExcelWorkbookProtectionSettings(
            false, true, false, Optional.of("book-secret"), Optional.of("review-secret")),
        protectionCommand.protection());

    assertInstanceOf(
        WorkbookSheetCommand.ClearWorkbookProtection.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookSelector.Current(), new WorkbookMutationAction.ClearWorkbookProtection()));

    WorkbookTabularCommand.SetAutofilter simpleAutofilter =
        assertInstanceOf(
            WorkbookTabularCommand.SetAutofilter.class,
            WorkbookCommandConverter.toCommand(
                new RangeSelector.ByRange("Budget", "A1:C4"),
                new StructuredMutationAction.SetAutofilter()));
    assertEquals(List.of(), simpleAutofilter.criteria());
    assertEquals(Optional.empty(), simpleAutofilter.sortState());
    assertEquals(List.of(), new StructuredMutationAction.SetAutofilter().criteria());

    WorkbookTabularCommand.SetAutofilter advancedAutofilter =
        assertInstanceOf(
            WorkbookTabularCommand.SetAutofilter.class,
            WorkbookCommandConverter.toCommand(
                new RangeSelector.ByRange("Budget", "A1:F9"), advancedAutofilterAction()));
    assertEquals(expectedAutofilterCriteria(), advancedAutofilter.criteria());
    assertEquals(Optional.of(expectedAutofilterSortState()), advancedAutofilter.sortState());

    WorkbookMetadataCommand.SetNamedRange namedRangeCommand =
        assertInstanceOf(
            WorkbookMetadataCommand.SetNamedRange.class,
            WorkbookCommandConverter.toCommand(
                new NamedRangeSelector.WorkbookScope("BudgetExpr"),
                new StructuredMutationAction.SetNamedRange(
                    "BudgetExpr",
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
            "sheet-secret");
    assertEquals(Optional.of("sheet-secret"), sheetProtection.password());
    org.junit.jupiter.api.Assertions.assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookMutationAction.SetSheetProtection(
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
    PictureDataInput pictureData =
        new PictureDataInput(ExcelPictureFormat.PNG, binary(PNG_PIXEL_BASE64));

    WorkbookDrawingCommand.SetPicture pictureCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetPicture.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName("Ops"),
                new DrawingMutationAction.SetPicture(
                    new PictureInput(
                        "OpsPicture", pictureData, anchor, Optional.of(text("Queue preview"))))));
    WorkbookDrawingCommand.SetSignatureLine signatureLineCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetSignatureLine.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName("Ops"),
                new DrawingMutationAction.SetSignatureLine(
                    new SignatureLineInput(
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
                new SheetSelector.ByName("Ops"),
                new DrawingMutationAction.SetShape(
                    new ShapeInput.SimpleShape(
                        "OpsShape", anchor, "rect", Optional.of(text("Queue"))))));
    WorkbookDrawingCommand.SetEmbeddedObject embeddedObjectCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetEmbeddedObject.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName("Ops"),
                new DrawingMutationAction.SetEmbeddedObject(
                    new EmbeddedObjectInput(
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
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new DrawingMutationAction.SetDrawingObjectAnchor(anchor)));
    WorkbookDrawingCommand.DeleteDrawingObject deleteCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.DeleteDrawingObject.class,
            WorkbookCommandConverter.toCommand(
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new DrawingMutationAction.DeleteDrawingObject()));

    assertEquals(
        new ExcelPictureDefinition(
            "OpsPicture",
            new ExcelBinaryData(java.util.Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
            ExcelPictureFormat.PNG,
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 4),
                new ExcelDrawingMarker(4, 6, 7, 8),
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
            Optional.of("Queue preview")),
        pictureCommand.picture());
    assertEquals(
        new ExcelSignatureLineDefinition(
            "OpsSignature",
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 4),
                new ExcelDrawingMarker(4, 6, 7, 8),
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
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 4),
                new ExcelDrawingMarker(4, 6, 7, 8),
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
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 3, 4),
                new ExcelDrawingMarker(4, 6, 7, 8),
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
                new RangeSelector.ByRange("Calc", "D2:D4"),
                new CellMutationAction.SetArrayFormula(
                    new ArrayFormulaInput(text("{=B2:B4*C2:C4}")))));
    WorkbookCellCommand.ClearArrayFormula clearArrayFormula =
        assertInstanceOf(
            WorkbookCellCommand.ClearArrayFormula.class,
            WorkbookCommandConverter.toCommand(
                new CellSelector.ByAddress("Calc", "D3"),
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
                    new CustomXmlImportInput(
                        new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                        text("<CORSO><NOME>Grid</NOME></CORSO>")))));

    assertEquals(
        new ExcelCustomXmlImportDefinition(
            new ExcelCustomXmlMappingLocator(1L, "CORSO_mapping"),
            "<CORSO><NOME>Grid</NOME></CORSO>"),
        command.mapping());
  }

  @Test
  void convertsPivotTableMutationOperations() {
    WorkbookTabularCommand.SetPivotTable setPivotTable =
        assertInstanceOf(
            WorkbookTabularCommand.SetPivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", "Report"),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableInput(
                        "Sales Pivot 2026",
                        "Report",
                        new PivotTableInput.Source.Range("Data", "A1:D5"),
                        new PivotTableInput.Anchor("C5"),
                        List.of("Region"),
                        List.of("Stage"),
                        List.of("Owner"),
                        List.of(
                            new PivotTableInput.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Amount",
                                Optional.of("#,##0.00")))))));
    WorkbookTabularCommand.SetPivotTable setPivotTableFromNamedRange =
        assertInstanceOf(
            WorkbookTabularCommand.SetPivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Named Source Pivot", "Report"),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableInput(
                        "Named Source Pivot",
                        "Report",
                        new PivotTableInput.Source.NamedRange("PivotSource"),
                        new PivotTableInput.Anchor("A3"),
                        List.of("Region"),
                        List.of(),
                        List.of(),
                        List.of(
                            new PivotTableInput.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                Optional.empty()))))));
    WorkbookTabularCommand.SetPivotTable setPivotTableFromTable =
        assertInstanceOf(
            WorkbookTabularCommand.SetPivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Table Source Pivot", "Report"),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableInput(
                        "Table Source Pivot",
                        "Report",
                        new PivotTableInput.Source.Table("SalesTable2026"),
                        new PivotTableInput.Anchor("G4"),
                        List.of("Region"),
                        List.of(),
                        List.of(),
                        List.of(
                            new PivotTableInput.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                Optional.empty()))))));
    WorkbookTabularCommand.DeletePivotTable deletePivotTable =
        assertInstanceOf(
            WorkbookTabularCommand.DeletePivotTable.class,
            WorkbookCommandConverter.toCommand(
                new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", "Report"),
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
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(1, 2, 3, 4),
            new DrawingMarkerInput(7, 12, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ChartSeriesInput firstSeries = chartSeries(new ChartTitleInput.Formula("B1"), "A2:A4", "B2:B4");
    ChartSeriesInput secondSeries =
        chartSeries(new ChartTitleInput.Text(text("Actual")), "ChartCategories", "ChartActual");

    WorkbookDrawingCommand.SetChart lineCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName("Ops"),
                new DrawingMutationAction.SetChart(
                    chartInput(
                        "TrendChart",
                        anchor,
                        new ChartTitleInput.Text(text("Trend")),
                        new ChartLegendInput.Hidden(),
                        ExcelChartDisplayBlanksAs.ZERO,
                        false,
                        new ChartPlotInput.Line(
                            true, ExcelChartGrouping.STANDARD, List.of(firstSeries))))));
    WorkbookDrawingCommand.SetChart pieCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName("Ops"),
                new DrawingMutationAction.SetChart(
                    chartInput(
                        "ShareChart",
                        anchor,
                        new ChartTitleInput.Formula("C1"),
                        null,
                        null,
                        null,
                        new ChartPlotInput.Pie(false, Optional.of(120), List.of(secondSeries))))));

    assertEquals("Ops", lineCommand.sheetName());
    ExcelChartDefinition lineChart = lineCommand.chart();
    ExcelChartDefinition.Line linePlot =
        assertInstanceOf(ExcelChartDefinition.Line.class, lineChart.plots().getFirst());
    assertEquals(new ExcelChartDefinition.Title.Text("Trend"), lineChart.title());
    assertEquals(new ExcelChartDefinition.Legend.Hidden(), lineChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.ZERO, lineChart.displayBlanksAs());
    assertFalse(lineChart.plotOnlyVisibleCells());
    assertTrue(linePlot.varyColors());
    assertEquals(
        chartDefinitionSeries(new ExcelChartDefinition.Title.Formula("B1"), "A2:A4", "B2:B4"),
        linePlot.series().getFirst());

    ExcelChartDefinition pieChart = pieCommand.chart();
    ExcelChartDefinition.Pie piePlot =
        assertInstanceOf(ExcelChartDefinition.Pie.class, pieChart.plots().getFirst());
    assertEquals(new ExcelChartDefinition.Title.Formula("C1"), pieChart.title());
    assertEquals(
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT), pieChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, pieChart.displayBlanksAs());
    assertTrue(pieChart.plotOnlyVisibleCells());
    assertFalse(piePlot.varyColors());
    assertEquals(Optional.of(120), piePlot.firstSliceAngle());
    assertEquals(
        chartDefinitionSeries(
            new ExcelChartDefinition.Title.Text("Actual"), "ChartCategories", "ChartActual"),
        piePlot.series().getFirst());
  }

  @Test
  void directChartHelperConversionsCoverStandaloneSwitchBranches() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(2, 3, 0, 0),
            new DrawingMarkerInput(9, 16, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
    ChartSeriesInput barSeries =
        chartSeries(new ChartTitleInput.None(), "Summary!$A$2:$A$4", "Summary!$B$2:$B$4");
    ChartInput barInput =
        chartInput(
            "OpsBar",
            anchor,
            new ChartTitleInput.Formula("Summary!$B$1"),
            new ChartLegendInput.Visible(ExcelChartLegendPosition.TOP),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            new ChartPlotInput.Bar(
                true,
                ExcelChartBarDirection.BAR,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                List.of(barSeries)));
    ChartInput lineInput =
        chartInput(
            "OpsLine",
            anchor,
            new ChartTitleInput.Text(text("Trend")),
            new ChartLegendInput.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            true,
            new ChartPlotInput.Line(
                false,
                ExcelChartGrouping.STANDARD,
                List.of(
                    chartSeries(
                        new ChartTitleInput.Formula("Summary!$C$1"),
                        "ChartCategories",
                        "ChartActual"))));
    ChartInput pieInput =
        chartInput(
            "OpsPie",
            anchor,
            new ChartTitleInput.Text(text("Share")),
            new ChartLegendInput.Visible(ExcelChartLegendPosition.LEFT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            new ChartPlotInput.Pie(
                true,
                Optional.of(90),
                List.of(
                    chartSeries(
                        new ChartTitleInput.Text(text("Actual")),
                        "ChartCategories",
                        "ChartActual"))));

    ExcelChartDefinition bar =
        WorkbookCommandDrawingInputConverter.toExcelChartDefinition(barInput);
    ExcelChartDefinition.Line linePlot =
        assertInstanceOf(
            ExcelChartDefinition.Line.class,
            WorkbookCommandDrawingInputConverter.toExcelChartDefinition(lineInput)
                .plots()
                .getFirst());
    ExcelChartDefinition.Pie piePlot =
        assertInstanceOf(
            ExcelChartDefinition.Pie.class,
            WorkbookCommandDrawingInputConverter.toExcelChartDefinition(pieInput)
                .plots()
                .getFirst());
    WorkbookDrawingCommand.SetChart barCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName("Ops"), new DrawingMutationAction.SetChart(barInput)));

    assertEquals(new ExcelChartDefinition.Title.Formula("Summary!$B$1"), bar.title());
    assertEquals(
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.TOP), bar.legend());
    assertEquals(ExcelChartDisplayBlanksAs.SPAN, bar.displayBlanksAs());
    assertFalse(bar.plotOnlyVisibleCells());
    ExcelChartDefinition.Bar barPlot =
        assertInstanceOf(ExcelChartDefinition.Bar.class, bar.plots().getFirst());
    assertTrue(barPlot.varyColors());
    assertEquals(ExcelChartBarDirection.BAR, barPlot.barDirection());
    assertEquals(ExcelChartBarGrouping.CLUSTERED, barPlot.grouping());
    assertEquals(
        chartDefinitionSeries(
            new ExcelChartDefinition.Title.None(), "Summary!$A$2:$A$4", "Summary!$B$2:$B$4"),
        barPlot.series().getFirst());
    ExcelChartDefinition line =
        WorkbookCommandDrawingInputConverter.toExcelChartDefinition(lineInput);
    assertEquals(new ExcelChartDefinition.Title.Text("Trend"), line.title());
    assertEquals(new ExcelChartDefinition.Legend.Hidden(), line.legend());
    assertEquals(
        new ExcelChartDefinition.Title.Formula("Summary!$C$1"),
        linePlot.series().getFirst().title());
    ExcelChartDefinition pie =
        WorkbookCommandDrawingInputConverter.toExcelChartDefinition(pieInput);
    assertEquals(new ExcelChartDefinition.Title.Text("Share"), pie.title());
    assertEquals(
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.LEFT), pie.legend());
    assertTrue(piePlot.varyColors());
    assertEquals(Optional.of(90), piePlot.firstSliceAngle());
    assertEquals("Ops", barCommand.sheetName());
    assertEquals(bar, barCommand.chart());
  }

  @Test
  void convertsAdvancedStyleAndConditionalFormattingPayloads() {
    ExcelCellStyle style =
        WorkbookCommandCellInputConverter.toExcelCellStyle(
            new CellStylePatchInput(
                Optional.empty(),
                Optional.empty(),
                Optional.of(
                    new CellFontInput(
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of(ColorInput.theme(2, 0.4d)),
                        Optional.of(true),
                        Optional.empty())),
                Optional.of(
                    CellFillInput.gradient(
                        CellGradientFillInput.path(
                            Optional.of(0.1d),
                            Optional.of(0.2d),
                            Optional.of(0.3d),
                            Optional.of(0.4d),
                            List.of(
                                new CellGradientStopInput(0.0d, ColorInput.rgb("#112233")),
                                new CellGradientStopInput(1.0d, ColorInput.theme(5, 0.2d)))))),
                Optional.of(
                    new CellBorderInput(
                        Optional.empty(),
                        Optional.ofNullable(new BorderSideInput(null, ColorInput.theme(1, 0.15d))),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())),
                Optional.empty()));

    assertEquals(
        Optional.of(
            new ExcelCellFont(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(ExcelColor.theme(2, 0.4d)),
                Optional.of(true),
                Optional.empty())),
        style.font());
    assertEquals(
        Optional.of(
            ExcelCellFill.gradient(
                ExcelGradientFill.path(
                    Optional.of(0.1d),
                    Optional.of(0.2d),
                    Optional.of(0.3d),
                    Optional.of(0.4d),
                    List.of(
                        new ExcelGradientStop(0.0d, ExcelColor.rgb("#112233")),
                        new ExcelGradientStop(1.0d, ExcelColor.theme(5, 0.2d)))))),
        style.fill());
    assertEquals(
        Optional.of(
            new ExcelBorder(
                Optional.empty(),
                Optional.ofNullable(
                    new ExcelBorderSide(Optional.empty(), Optional.of(ExcelColor.theme(1, 0.15d)))),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())),
        style.border());

    assertEquals(
        new ExcelConditionalFormattingRule.ColorScaleRule(
            List.of(
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.MIN, null, null),
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.MAX, null, null)),
            List.of(ExcelColor.rgb("#112233"), ExcelColor.rgb("#AABBCC")),
            true),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.ColorScaleRule(
                true,
                List.of(
                    new ConditionalFormattingThresholdInput.Min(),
                    new ConditionalFormattingThresholdInput.Max()),
                List.of(ColorInput.rgb("#112233"), ColorInput.rgb("#AABBCC")))));

    assertEquals(
        new ExcelConditionalFormattingRule.DataBarRule(
            ExcelColor.theme(4, 0.25d),
            true,
            10,
            90,
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.PERCENTILE, null, 90.0d),
            false),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                ColorInput.theme(4, 0.25d),
                true,
                10,
                90,
                new ConditionalFormattingThresholdInput.Numeric(0.0d),
                new ConditionalFormattingThresholdInput.Percentile(90.0d))));

    ExcelConditionalFormattingRule.ColorScaleRule formulaThresholdRule =
        assertInstanceOf(
            ExcelConditionalFormattingRule.ColorScaleRule.class,
            WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
                new ConditionalFormattingRuleInput.ColorScaleRule(
                    false,
                    List.of(
                        new ConditionalFormattingThresholdInput.Formula("A1"),
                        new ConditionalFormattingThresholdInput.Percent(100.0d)),
                    List.of(ColorInput.rgb("#112233"), ColorInput.rgb("#AABBCC")))));
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
            new ConditionalFormattingRuleInput.IconSetRule(
                true,
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                true,
                List.of(
                    new ConditionalFormattingThresholdInput.Percent(0.0d),
                    new ConditionalFormattingThresholdInput.Percent(33.0d),
                    new ConditionalFormattingThresholdInput.Percent(67.0d)))));

    assertEquals(
        new ExcelConditionalFormattingRule.Top10Rule(7, true, false, true, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.Top10Rule(true, 7, true, false, Optional.empty())));
  }

  @Test
  void executorContextHelpersReturnNoFalseMetadataForWorkbookProtectionOperations() {
    WorkbookMutationAction.SetWorkbookProtection setProtection =
        new WorkbookMutationAction.SetWorkbookProtection(
            new WorkbookProtectionInput(
                true, false, true, Optional.of("book-secret"), Optional.empty()));
    WorkbookMutationAction.ClearWorkbookProtection clearProtection =
        new WorkbookMutationAction.ClearWorkbookProtection();

    assertEquals("SET_WORKBOOK_PROTECTION", setProtection.actionType());
    assertEquals("CLEAR_WORKBOOK_PROTECTION", clearProtection.actionType());
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.formulaFor(
            new MutationStep(
                "step-01-set-workbook-protection", new WorkbookSelector.Current(), setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.sheetNameFor(
            new MutationStep(
                "step-01-set-workbook-protection", new WorkbookSelector.Current(), setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.addressFor(
            new MutationStep(
                "step-01-set-workbook-protection", new WorkbookSelector.Current(), setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.rangeFor(
            new MutationStep(
                "step-01-set-workbook-protection", new WorkbookSelector.Current(), setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new MutationStep(
                "step-01-set-workbook-protection", new WorkbookSelector.Current(), setProtection),
            new IllegalStateException()));

    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.formulaFor(
            new MutationStep(
                "step-02-clear-workbook-protection",
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.sheetNameFor(
            new MutationStep(
                "step-02-clear-workbook-protection",
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.addressFor(
            new MutationStep(
                "step-02-clear-workbook-protection",
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.rangeFor(
            new MutationStep(
                "step-02-clear-workbook-protection",
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new MutationStep(
                "step-02-clear-workbook-protection",
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
  }

  private static StructuredMutationAction.SetAutofilter advancedAutofilterAction() {
    return new StructuredMutationAction.SetAutofilter(
        List.of(
            AutofilterFilterColumnInput.visibleButton(
                0L, new AutofilterFilterCriterionInput.Values(List.of("Queued", "Ready"), true)),
            new AutofilterFilterColumnInput(
                1L,
                false,
                new AutofilterFilterCriterionInput.Custom(
                    true,
                    List.of(
                        new AutofilterFilterCriterionInput.CustomConditionInput(
                            "greaterThan", "5")))),
            new AutofilterFilterColumnInput(
                2L,
                true,
                new AutofilterFilterCriterionInput.Dynamic(
                    "TODAY", java.util.Optional.of(1.0d), java.util.Optional.of(2.0d))),
            new AutofilterFilterColumnInput(
                3L, true, new AutofilterFilterCriterionInput.Top10(10, true, false)),
            new AutofilterFilterColumnInput(
                4L,
                true,
                new AutofilterFilterCriterionInput.Color(false, ColorInput.theme(3, 0.25d))),
            new AutofilterFilterColumnInput(
                5L, true, new AutofilterFilterCriterionInput.Icon("3TrafficLights1", 2))),
        AutofilterSortStateInput.withoutSortMethod(
            "A2:F9",
            false,
            true,
            List.of(
                new AutofilterSortConditionInput.CellColor(
                    "B2:B9", true, ColorInput.rgb("#AABBCC")),
                new AutofilterSortConditionInput.FontColor(
                    "C2:C9", false, ColorInput.theme(4, 0.15d)),
                new AutofilterSortConditionInput.Icon("C2:C9", false, 2))));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  private static BinarySourceInput binary(String value) {
    return BinarySourceInput.inlineBase64(value);
  }

  private static ChartInput chartInput(
      String name,
      DrawingAnchorInput.TwoCell anchor,
      ChartTitleInput title,
      ChartLegendInput legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      ChartPlotInput plot) {
    return new ChartInput(
        name,
        anchor,
        title == null ? new ChartTitleInput.None() : title,
        legend == null ? new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT) : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        List.of(plot));
  }

  private static ChartSeriesInput chartSeries(
      ChartTitleInput title, String categoriesFormula, String valuesFormula) {
    return new ChartSeriesInput(
        title,
        new ChartDataSourceInput.Reference(categoriesFormula),
        new ChartDataSourceInput.Reference(valuesFormula),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static ExcelChartDefinition.Series chartDefinitionSeries(
      ExcelChartDefinition.Title title, String categoriesFormula, String valuesFormula) {
    return new ExcelChartDefinition.Series(
        title,
        new ExcelChartDefinition.DataSource.Reference(categoriesFormula),
        new ExcelChartDefinition.DataSource.Reference(valuesFormula),
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
            4L, true, new ExcelAutofilterFilterCriterion.Color(false, ExcelColor.theme(3, 0.25d))),
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
            new ExcelAutofilterSortCondition.CellColor("B2:B9", true, ExcelColor.rgb("#AABBCC")),
            new ExcelAutofilterSortCondition.FontColor("C2:C9", false, ExcelColor.theme(4, 0.15d)),
            new ExcelAutofilterSortCondition.Icon("C2:C9", false, 2)));
  }
}
