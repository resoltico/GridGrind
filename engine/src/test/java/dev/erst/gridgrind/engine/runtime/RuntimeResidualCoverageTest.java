package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused closure tests for runtime seams left uncovered by broader end-to-end suites. */
class RuntimeResidualCoverageTest {

  @Test
  void converterNullBranchesAndResolvedPromptFieldsStayTyped() {
    assertTrue(WorkbookCommandStructuredInputConverter.toResolvedValidationPrompt(null).isEmpty());
    assertTrue(
        WorkbookCommandStructuredInputConverter.toResolvedValidationErrorAlert(null).isEmpty());

    Optional<ResolvedValidationPrompt> prompt =
        WorkbookCommandStructuredInputConverter.toResolvedValidationPrompt(
            new DataValidationPromptSpec(
                TextSourceSpec.inline("Prompt title"), TextSourceSpec.inline("Prompt text"), true));
    Optional<ResolvedValidationErrorAlert> alert =
        WorkbookCommandStructuredInputConverter.toResolvedValidationErrorAlert(
            new DataValidationErrorAlertSpec(
                ExcelDataValidationErrorStyle.WARNING,
                TextSourceSpec.inline("Alert title"),
                TextSourceSpec.inline("Alert text"),
                true));

    assertEquals("Prompt title", prompt.orElseThrow().title());
    assertEquals(ExcelDataValidationErrorStyle.WARNING, alert.orElseThrow().style());
  }

  @Test
  void sourceBackedRequirementsCoverResidualDispatchBranches() {
    List<CellMutationAction> noInputCellActions =
        List.of(
            new CellMutationAction.ClearRange(),
            new CellMutationAction.SetArrayFormula(
                new ArrayFormulaSpec(TextSourceSpec.inline("{=SUM(A1:A2)}"))),
            new CellMutationAction.ClearArrayFormula(),
            new CellMutationAction.SetHyperlink(new HyperlinkTarget.Url("https://example.com")),
            new CellMutationAction.ClearHyperlink(),
            new CellMutationAction.ClearComment(),
            new CellMutationAction.ApplyStyle(
                new CellStylePatch(
                    Optional.of("0.00"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())));
    for (CellMutationAction action : noInputCellActions) {
      assertFalse(SourceBackedInputRequirements.requiresStandardInput(action));
    }

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellMutationAction.SetCell(
                new CellValueSpec.Text(TextSourceSpec.standardInput()))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                    List.of(List.of(new CellValueSpec.Text(TextSourceSpec.standardInput())))))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellMutationAction.SetComment(
                CommentSpec.plain(TextSourceSpec.standardInput(), "Ada", false))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellMutationAction.AppendRow(
                new dev.erst.gridgrind.core.model.value.CellRowSpec.Typed(
                    List.of(new CellValueSpec.Formula(TextSourceSpec.standardInput()))))));

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetPicture(
                new PictureSpec(
                    "Banner",
                    pictureData(BinarySourceSpec.standardInput()),
                    anchor(),
                    Optional.of(TextSourceSpec.inline("Preview"))))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new PictureSpec(
                "Described Banner",
                pictureData(BinarySourceSpec.file("banner.png")),
                anchor(),
                Optional.of(TextSourceSpec.standardInput()))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new PictureSpec(
                "Inline Description",
                pictureData(BinarySourceSpec.file("banner.png")),
                anchor(),
                Optional.of(TextSourceSpec.inline("Preview")))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetSignatureLine(
                signatureLine(Optional.of(pictureData(BinarySourceSpec.standardInput()))))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetChart(
                chartWithTitle(new ChartTitleSpec.Text(TextSourceSpec.standardInput())))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetShape(
                new ShapeSpec.SimpleShape(
                    "Callout", anchor(), "rect", Optional.of(TextSourceSpec.standardInput())))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectSpec(
                    "Attachment",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceSpec.standardInput(),
                    pictureData(BinarySourceSpec.file("preview.png")),
                    anchor()))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetDrawingObjectAnchor(anchor())));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.DeleteDrawingObject()));

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.ImportCustomXmlMapping(
                new CustomXmlImportSpec(
                    new CustomXmlMappingLocator(
                        java.util.Optional.empty(), java.util.Optional.of("root/item")),
                    TextSourceSpec.standardInput()))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.SetPivotTable(pivotTable())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.SetDataValidation(validation())));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.ClearDataValidations()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.SetConditionalFormatting(
                new ConditionalFormattingSpec(
                    List.of(
                        new ConditionalFormattingRuleSpec.FormulaRule(
                            "TRUE", true, Optional.empty()))))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.ClearConditionalFormatting()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.SetAutofilter()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.ClearAutofilter()));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.SetTable(table(TextSourceSpec.standardInput()))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.DeleteTable()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.DeletePivotTable()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.SetNamedRange(
                new DefinedName("Budget"),
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.formula("SUM(Budget!A1:A3)"))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.DeleteNamedRange()));

    List<WorkbookMutationAction> noInputWorkbookActions =
        List.of(
            new WorkbookMutationAction.EnsureSheet(),
            new WorkbookMutationAction.RenameSheet(new SheetName("Renamed")),
            new WorkbookMutationAction.DeleteSheet(),
            new WorkbookMutationAction.MoveSheet(1),
            new WorkbookMutationAction.CopySheet(new SheetName("Budget Copy")),
            new WorkbookMutationAction.SetActiveSheet(),
            new WorkbookMutationAction.SetSelectedSheets(),
            new WorkbookMutationAction.SetSheetVisibility(
                dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility.HIDDEN),
            new WorkbookMutationAction.SetSheetProtection(
                new SheetProtectionSettings(
                    false, false, false, false, false, false, false, false, false, false, false,
                    false, false, false, false)),
            new WorkbookMutationAction.ClearSheetProtection(),
            new WorkbookMutationAction.SetWorkbookProtection(
                new WorkbookProtectionSpec(true, false, false, Optional.empty(), Optional.empty())),
            new WorkbookMutationAction.ClearWorkbookProtection(),
            new WorkbookMutationAction.MergeCells(),
            new WorkbookMutationAction.UnmergeCells(),
            new WorkbookMutationAction.SetColumnWidth(12.0d),
            new WorkbookMutationAction.SetRowHeight(14.0d),
            new WorkbookMutationAction.InsertRows(),
            new WorkbookMutationAction.DeleteRows(),
            new WorkbookMutationAction.ShiftRows(1),
            new WorkbookMutationAction.InsertColumns(),
            new WorkbookMutationAction.DeleteColumns(),
            new WorkbookMutationAction.ShiftColumns(1),
            new WorkbookMutationAction.SetRowVisibility(true),
            new WorkbookMutationAction.SetColumnVisibility(true),
            WorkbookMutationAction.GroupRows.expanded(),
            new WorkbookMutationAction.UngroupRows(),
            WorkbookMutationAction.GroupColumns.expanded(),
            new WorkbookMutationAction.UngroupColumns(),
            new WorkbookMutationAction.SetSheetPane(new PaneSpec.Frozen(1, 0, 1, 0)),
            new WorkbookMutationAction.SetSheetZoom(125),
            new WorkbookMutationAction.SetSheetPresentation(SheetPresentationSpec.defaults()),
            new WorkbookMutationAction.ClearPrintLayout(),
            new WorkbookMutationAction.AutoSizeColumns());
    for (WorkbookMutationAction action : noInputWorkbookActions) {
      assertFalse(SourceBackedInputRequirements.requiresStandardInput(action));
    }
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new WorkbookMutationAction.SetPrintLayout(
                new PrintLayoutSpec(
                    PrintLayoutSpec.defaults().printArea(),
                    ExcelPrintOrientation.PORTRAIT,
                    PrintLayoutSpec.defaults().scaling(),
                    PrintLayoutSpec.defaults().repeatingRows(),
                    PrintLayoutSpec.defaults().repeatingColumns(),
                    new HeaderFooterTextSpec(
                        TextSourceSpec.standardInput(),
                        TextSourceSpec.inline(""),
                        TextSourceSpec.inline("")),
                    HeaderFooterTextSpec.blank(),
                    PrintLayoutSpec.defaults().setup()))));

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellValueSpec.RichText(
                List.of(new RichTextRunSpec(TextSourceSpec.standardInput(), Optional.empty())))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartTitleSpec.Text(TextSourceSpec.standardInput())));
    for (ChartPlotSpec plot : allPlotKindsWithStandardInputSeries()) {
      assertTrue(SourceBackedInputRequirements.requiresStandardInput(plot));
    }
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new ShapeSpec.SimpleShape(
                "StandardTextShape",
                anchor(),
                "rect",
                Optional.of(TextSourceSpec.standardInput()))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ShapeSpec.Connector("Link", anchor())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new EmbeddedObjectSpec(
                "Embed",
                "Payload",
                "payload.bin",
                "open",
                BinarySourceSpec.standardInput(),
                pictureData(BinarySourceSpec.file("preview.png")),
                anchor())));
    assertTrue(SourceBackedInputRequirements.requiresStandardInput(validation()));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(table(TextSourceSpec.standardInput())));
    assertFalse(SourceBackedInputRequirements.requiresStandardInput(PrintLayoutSpec.defaults()));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new PrintLayoutSpec(
                PrintLayoutSpec.defaults().printArea(),
                ExcelPrintOrientation.PORTRAIT,
                PrintLayoutSpec.defaults().scaling(),
                PrintLayoutSpec.defaults().repeatingRows(),
                PrintLayoutSpec.defaults().repeatingColumns(),
                new HeaderFooterTextSpec(
                    TextSourceSpec.standardInput(),
                    TextSourceSpec.inline(""),
                    TextSourceSpec.inline("")),
                HeaderFooterTextSpec.blank(),
                PrintLayoutSpec.defaults().setup())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CustomXmlImportSpec(
                new CustomXmlMappingLocator(
                    java.util.Optional.empty(), java.util.Optional.of("root/item")),
                TextSourceSpec.standardInput())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new HeaderFooterTextSpec(
                TextSourceSpec.inline(""),
                TextSourceSpec.standardInput(),
                TextSourceSpec.inline(""))));
    assertTrue(SourceBackedInputRequirements.requiresStandardInput(TextSourceSpec.standardInput()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(TextSourceSpec.inline("inline")));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(BinarySourceSpec.standardInput()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(BinarySourceSpec.file("payload.bin")));

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByName(new TableName("BudgetTable")),
                    "Owner",
                    new CellValueSpec.Text(TextSourceSpec.standardInput())),
                "Amount")));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName(new TableName("BudgetTable")),
                "Owner",
                new CellValueSpec.Text(TextSourceSpec.standardInput()))));
    List<Selector> noInputSelectors =
        List.of(
            new TableRowSelector.AllRows(new TableSelector.ByName(new TableName("BudgetTable"))),
            new TableRowSelector.ByIndex(new TableSelector.ByName(new TableName("BudgetTable")), 1),
            new WorkbookSelector.Current(),
            new SheetSelector.ByName(new SheetName("Budget")),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
            new RowBandSelector.Span(new SheetName("Budget"), 0, 1),
            new ColumnBandSelector.Span(new SheetName("Budget"), 0, 1),
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "Banner"),
            new ChartSelector.ByName(new SheetName("Budget"), "LoadChart"),
            new TableSelector.ByName(new TableName("BudgetTable")),
            new PivotTableSelector.ByName("Sales Pivot"),
            new NamedRangeSelector.ByName(new DefinedName("BudgetRange")));
    for (Selector selector : noInputSelectors) {
      assertFalse(SourceBackedInputRequirements.requiresStandardInput(selector));
    }
  }

  private static DrawingAnchorSpec.TwoCell anchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(0, 0),
        new DrawingMarkerSpec(2, 2),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static PictureDataSpec pictureData(BinarySourceSpec source) {
    return new PictureDataSpec(ExcelPictureFormat.PNG, source);
  }

  private static SignatureLineSpec signatureLine(Optional<PictureDataSpec> plainSignature) {
    return new SignatureLineSpec(
        "Approval Signature",
        anchor(),
        true,
        Optional.of("Review before signing"),
        Optional.of("Ada Lovelace"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.of("invalid"),
        plainSignature);
  }

  private static ChartSpec chartWithTitle(ChartTitleSpec title) {
    return new ChartSpec(
        "Load Chart",
        anchor(),
        title,
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        List.of(
            new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(series(title)))));
  }

  private static List<ChartPlotSpec> allPlotKindsWithStandardInputSeries() {
    ChartSeriesSpec series = series(new ChartTitleSpec.Text(TextSourceSpec.standardInput()));
    return List.of(
        new ChartPlotSpec.Area(false, ExcelChartGrouping.STANDARD, List.of(series)),
        new ChartPlotSpec.Area3D(
            false, ExcelChartGrouping.STANDARD, Optional.empty(), List.of(series)),
        new ChartPlotSpec.Bar(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.empty(),
            Optional.empty(),
            List.of(series)),
        new ChartPlotSpec.Bar3D(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            List.of(series)),
        new ChartPlotSpec.Doughnut(false, Optional.empty(), Optional.of(50), List.of(series)),
        new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, List.of(series)),
        new ChartPlotSpec.Line3D(
            false, ExcelChartGrouping.STANDARD, Optional.empty(), List.of(series)),
        new ChartPlotSpec.Pie(false, Optional.empty(), List.of(series)),
        new ChartPlotSpec.Pie3D(false, List.of(series)),
        new ChartPlotSpec.Radar(false, ExcelChartRadarStyle.STANDARD, List.of(series)),
        new ChartPlotSpec.Scatter(false, ExcelChartScatterStyle.LINE, List.of(series)),
        new ChartPlotSpec.Surface(false, false, List.of(series)),
        new ChartPlotSpec.Surface3D(false, false, List.of(series)));
  }

  private static ChartSeriesSpec series(ChartTitleSpec title) {
    return new ChartSeriesSpec(
        title,
        new ChartDataSourceSpec.StringLiteral(List.of("Jan", "Feb")),
        new ChartDataSourceSpec.NumericLiteral(List.of(10d, 12d)),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static DataValidationSpec validation() {
    return new DataValidationSpec(
        new DataValidationRuleSpec.WholeNumber(
            ExcelComparisonOperator.BETWEEN, "1", Optional.of("10")),
        false,
        false,
        Optional.of(
            new DataValidationPromptSpec(
                TextSourceSpec.standardInput(), TextSourceSpec.inline("Prompt"), true)),
        Optional.of(
            new DataValidationErrorAlertSpec(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceSpec.inline("Alert"),
                TextSourceSpec.standardInput(),
                true)));
  }

  private static PivotTableSpec pivotTable() {
    return new PivotTableSpec(
        "Sales Pivot 2026",
        new SheetName("Report"),
        new PivotTableSpec.PivotSourceSpec.Range(new SheetName("Data"), new RangeRef("A1:D5")),
        new PivotTableSpec.Anchor(new CellAddress("C5")),
        List.of("Region"),
        List.of("Stage"),
        List.of(),
        List.of(
            new PivotTableSpec.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Amount", Optional.empty())));
  }

  private static TableSpec table(TextSourceSpec comment) {
    return new TableSpec(
        new TableName("BudgetTable"),
        new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
        new RangeRef("A1:C4"),
        false,
        false,
        new TableStyleSpec.None(),
        comment,
        false,
        false,
        false,
        "",
        "",
        "",
        List.of());
  }
}
