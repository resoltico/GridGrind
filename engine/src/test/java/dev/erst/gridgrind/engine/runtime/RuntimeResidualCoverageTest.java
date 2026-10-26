package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceInput;
import dev.erst.gridgrind.core.model.chart.ChartInput;
import dev.erst.gridgrind.core.model.chart.ChartLegendInput;
import dev.erst.gridgrind.core.model.chart.ChartPlotInput;
import dev.erst.gridgrind.core.model.chart.ChartSeriesInput;
import dev.erst.gridgrind.core.model.chart.ChartTitleInput;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportInput;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.CommentInput;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorInput;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerInput;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectInput;
import dev.erst.gridgrind.core.model.drawing.PictureDataInput;
import dev.erst.gridgrind.core.model.drawing.PictureInput;
import dev.erst.gridgrind.core.model.drawing.ShapeInput;
import dev.erst.gridgrind.core.model.drawing.SignatureLineInput;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaInput;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.PaneInput;
import dev.erst.gridgrind.core.model.layout.SheetPresentationInput;
import dev.erst.gridgrind.core.model.pivot.PivotTableInput;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextInput;
import dev.erst.gridgrind.core.model.print.PrintLayoutInput;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionInput;
import dev.erst.gridgrind.core.model.style.BorderSideInput;
import dev.erst.gridgrind.core.model.style.CellBorderInput;
import dev.erst.gridgrind.core.model.style.CellStylePatchInput;
import dev.erst.gridgrind.core.model.style.ColorInput;
import dev.erst.gridgrind.core.model.style.RichTextRunInput;
import dev.erst.gridgrind.core.model.table.TableInput;
import dev.erst.gridgrind.core.model.table.TableStyleInput;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingDefinitionInput;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertInput;
import dev.erst.gridgrind.core.model.validation.DataValidationInput;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptInput;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleInput;
import dev.erst.gridgrind.core.model.value.CellInput;
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
import dev.erst.gridgrind.core.source.BinarySourceInput;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
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
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationPrompt;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused closure tests for runtime seams left uncovered by broader end-to-end suites. */
class RuntimeResidualCoverageTest {
  @Test
  void secretBearingProtectionConversionsRequireExecutionBindings() {
    WorkbookProtectionInput workbookProtection =
        new WorkbookProtectionInput(
            true,
            false,
            false,
            Optional.of(new SecretReference("workbook-password")),
            Optional.empty());
    SheetProtectionSettings sheetProtection =
        new SheetProtectionSettings(
            false, false, false, false, false, false, false, false, false, false, false, false,
            false, false, false);

    assertThrows(
        IllegalStateException.class,
        () ->
            WorkbookCommandLayoutInputConverter.toExcelWorkbookProtectionSettings(
                workbookProtection, null));
    assertThrows(
        IllegalStateException.class,
        () ->
            WorkbookCommandWorkbookMutationConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.SetSheetProtection(
                    sheetProtection, Optional.of(new SecretReference("sheet-password"))),
                null));
  }

  @Test
  void converterNullBranchesAndColorReportVariantsStayTyped() {
    assertTrue(WorkbookCommandCellInputConverter.toExcelFontHeight(null).isEmpty());
    assertTrue(WorkbookCommandCellInputConverter.toExcelBorder(null).isEmpty());
    assertTrue(WorkbookCommandCellInputConverter.toExcelBorderSide(null).isEmpty());
    assertTrue(WorkbookCommandCellInputConverter.toExcelColor(null).isEmpty());
    assertTrue(WorkbookCommandStructuredInputConverter.toExcelDataValidationPrompt(null).isEmpty());
    assertTrue(
        WorkbookCommandStructuredInputConverter.toExcelDataValidationErrorAlert(null).isEmpty());
    assertTrue(WorkbookCommandStructuredInputConverter.toExcelAutofilterSortState(null).isEmpty());

    Optional<ExcelFontHeight> fontHeight =
        WorkbookCommandCellInputConverter.toExcelFontHeight(
            new dev.erst.gridgrind.core.model.style.FontHeightInput.Points(
                BigDecimal.valueOf(12.5d)));
    Optional<ExcelBorder> border =
        WorkbookCommandCellInputConverter.toExcelBorder(
            new CellBorderInput(
                Optional.of(
                    new BorderSideInput(
                        ExcelBorderStyle.THIN, ColorInput.rgb(new RgbHex("#112233")))),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    Optional<ExcelBorderSide> borderSide =
        WorkbookCommandCellInputConverter.toExcelBorderSide(
            new BorderSideInput(Optional.of(ExcelBorderStyle.DOUBLE), Optional.empty()));
    Optional<ExcelColor> color =
        WorkbookCommandCellInputConverter.toExcelColor(ColorInput.indexed(6, 0.25d));
    Optional<ExcelDataValidationPrompt> prompt =
        WorkbookCommandStructuredInputConverter.toExcelDataValidationPrompt(
            new DataValidationPromptInput(
                TextSourceInput.inline("Prompt title"),
                TextSourceInput.inline("Prompt text"),
                true));
    Optional<ExcelDataValidationErrorAlert> alert =
        WorkbookCommandStructuredInputConverter.toExcelDataValidationErrorAlert(
            new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.WARNING,
                TextSourceInput.inline("Alert title"),
                TextSourceInput.inline("Alert text"),
                true));

    assertEquals(Optional.of(new ExcelFontHeight(250)), fontHeight);
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN), border.orElseThrow().all().orElseThrow().style());
    assertEquals(Optional.of(ExcelBorderStyle.DOUBLE), borderSide.orElseThrow().style());
    assertEquals(Optional.of(0.25d), color.orElseThrow().tint());
    assertEquals("Prompt title", prompt.orElseThrow().title());
    assertEquals(ExcelDataValidationErrorStyle.WARNING, alert.orElseThrow().style());

    assertEquals(
        Optional.empty(),
        InspectionResultCellReportSupport.toCellColorReport((ExcelColorSnapshot) null));
    assertEquals(
        Optional.of(
            dev.erst.gridgrind.core.model.style.CellColorReport.rgb(new RgbHex("#112233"), 0.25d)),
        InspectionResultCellReportSupport.toCellColorReport(
            ExcelColorSnapshot.rgb("#112233", 0.25d)));
    assertEquals(
        Optional.of(dev.erst.gridgrind.core.model.style.CellColorReport.theme(3)),
        InspectionResultCellReportSupport.toCellColorReport(ExcelColorSnapshot.theme(3)));
    assertEquals(
        Optional.of(dev.erst.gridgrind.core.model.style.CellColorReport.indexed(7, -0.5d)),
        InspectionResultCellReportSupport.toCellColorReport(ExcelColorSnapshot.indexed(7, -0.5d)));
  }

  @Test
  void sourceBackedRequirementsCoverResidualDispatchBranches() {
    List<CellMutationAction> noInputCellActions =
        List.of(
            new CellMutationAction.ClearRange(),
            new CellMutationAction.SetArrayFormula(
                new ArrayFormulaInput(TextSourceInput.inline("{=SUM(A1:A2)}"))),
            new CellMutationAction.ClearArrayFormula(),
            new CellMutationAction.SetHyperlink(new HyperlinkTarget.Url("https://example.com")),
            new CellMutationAction.ClearHyperlink(),
            new CellMutationAction.ClearComment(),
            new CellMutationAction.ApplyStyle(
                new CellStylePatchInput(
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
            new CellMutationAction.SetCell(new CellInput.Text(TextSourceInput.standardInput()))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.core.model.value.CellGridInput.Typed(
                    List.of(List.of(new CellInput.Text(TextSourceInput.standardInput())))))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellMutationAction.SetComment(
                CommentInput.plain(TextSourceInput.standardInput(), "Ada", false))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellMutationAction.AppendRow(
                new dev.erst.gridgrind.core.model.value.CellRowInput.Typed(
                    List.of(new CellInput.Formula(TextSourceInput.standardInput()))))));

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetPicture(
                new PictureInput(
                    "Banner",
                    pictureData(BinarySourceInput.standardInput()),
                    anchor(),
                    Optional.of(TextSourceInput.inline("Preview"))))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new PictureInput(
                "Described Banner",
                pictureData(BinarySourceInput.file("banner.png")),
                anchor(),
                Optional.of(TextSourceInput.standardInput()))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new PictureInput(
                "Inline Description",
                pictureData(BinarySourceInput.file("banner.png")),
                anchor(),
                Optional.of(TextSourceInput.inline("Preview")))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetSignatureLine(
                signatureLine(Optional.of(pictureData(BinarySourceInput.standardInput()))))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetChart(
                chartWithTitle(new ChartTitleInput.Text(TextSourceInput.standardInput())))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetShape(
                new ShapeInput.SimpleShape(
                    "Callout", anchor(), "rect", Optional.of(TextSourceInput.standardInput())))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Attachment",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.standardInput(),
                    pictureData(BinarySourceInput.file("preview.png")),
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
                new CustomXmlImportInput(
                    new CustomXmlMappingLocator(null, "root/item"),
                    TextSourceInput.standardInput()))));
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
                new ConditionalFormattingDefinitionInput(
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
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
            new StructuredMutationAction.SetTable(table(TextSourceInput.standardInput()))));
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
                new WorkbookProtectionInput(
                    true, false, false, Optional.empty(), Optional.empty())),
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
            new WorkbookMutationAction.SetSheetPane(new PaneInput.Frozen(1, 0, 1, 0)),
            new WorkbookMutationAction.SetSheetZoom(125),
            new WorkbookMutationAction.SetSheetPresentation(SheetPresentationInput.defaults()),
            new WorkbookMutationAction.ClearPrintLayout(),
            new WorkbookMutationAction.AutoSizeColumns());
    for (WorkbookMutationAction action : noInputWorkbookActions) {
      assertFalse(SourceBackedInputRequirements.requiresStandardInput(action));
    }
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new WorkbookMutationAction.SetPrintLayout(
                new PrintLayoutInput(
                    PrintLayoutInput.defaults().printArea(),
                    ExcelPrintOrientation.PORTRAIT,
                    PrintLayoutInput.defaults().scaling(),
                    PrintLayoutInput.defaults().repeatingRows(),
                    PrintLayoutInput.defaults().repeatingColumns(),
                    new HeaderFooterTextInput(
                        TextSourceInput.standardInput(),
                        TextSourceInput.inline(""),
                        TextSourceInput.inline("")),
                    HeaderFooterTextInput.blank(),
                    PrintLayoutInput.defaults().setup()))));

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellInput.RichText(
                List.of(new RichTextRunInput(TextSourceInput.standardInput(), Optional.empty())))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartTitleInput.Text(TextSourceInput.standardInput())));
    for (ChartPlotInput plot : allPlotKindsWithStandardInputSeries()) {
      assertTrue(SourceBackedInputRequirements.requiresStandardInput(plot));
    }
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new ShapeInput.SimpleShape(
                "StandardTextShape",
                anchor(),
                "rect",
                Optional.of(TextSourceInput.standardInput()))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ShapeInput.Connector("Link", anchor())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new EmbeddedObjectInput(
                "Embed",
                "Payload",
                "payload.bin",
                "open",
                BinarySourceInput.standardInput(),
                pictureData(BinarySourceInput.file("preview.png")),
                anchor())));
    assertTrue(SourceBackedInputRequirements.requiresStandardInput(validation()));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            table(TextSourceInput.standardInput())));
    assertFalse(SourceBackedInputRequirements.requiresStandardInput(PrintLayoutInput.defaults()));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new PrintLayoutInput(
                PrintLayoutInput.defaults().printArea(),
                ExcelPrintOrientation.PORTRAIT,
                PrintLayoutInput.defaults().scaling(),
                PrintLayoutInput.defaults().repeatingRows(),
                PrintLayoutInput.defaults().repeatingColumns(),
                new HeaderFooterTextInput(
                    TextSourceInput.standardInput(),
                    TextSourceInput.inline(""),
                    TextSourceInput.inline("")),
                HeaderFooterTextInput.blank(),
                PrintLayoutInput.defaults().setup())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CustomXmlImportInput(
                new CustomXmlMappingLocator(null, "root/item"), TextSourceInput.standardInput())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new HeaderFooterTextInput(
                TextSourceInput.inline(""),
                TextSourceInput.standardInput(),
                TextSourceInput.inline(""))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(TextSourceInput.standardInput()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(TextSourceInput.inline("inline")));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(BinarySourceInput.standardInput()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(BinarySourceInput.file("payload.bin")));

    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByName(new TableName("BudgetTable")),
                    "Owner",
                    new CellInput.Text(TextSourceInput.standardInput())),
                "Amount")));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName(new TableName("BudgetTable")),
                "Owner",
                new CellInput.Text(TextSourceInput.standardInput()))));
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

  private static DrawingAnchorInput.TwoCell anchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(0, 0),
        new DrawingMarkerInput(2, 2),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static PictureDataInput pictureData(BinarySourceInput source) {
    return new PictureDataInput(ExcelPictureFormat.PNG, source);
  }

  private static SignatureLineInput signatureLine(Optional<PictureDataInput> plainSignature) {
    return new SignatureLineInput(
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

  private static ChartInput chartWithTitle(ChartTitleInput title) {
    return new ChartInput(
        "Load Chart",
        anchor(),
        title,
        new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        List.of(
            new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(series(title)))));
  }

  private static List<ChartPlotInput> allPlotKindsWithStandardInputSeries() {
    ChartSeriesInput series = series(new ChartTitleInput.Text(TextSourceInput.standardInput()));
    return List.of(
        new ChartPlotInput.Area(false, ExcelChartGrouping.STANDARD, List.of(series)),
        new ChartPlotInput.Area3D(
            false, ExcelChartGrouping.STANDARD, Optional.empty(), List.of(series)),
        new ChartPlotInput.Bar(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.empty(),
            Optional.empty(),
            List.of(series)),
        new ChartPlotInput.Bar3D(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            List.of(series)),
        new ChartPlotInput.Doughnut(false, Optional.empty(), Optional.of(50), List.of(series)),
        new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, List.of(series)),
        new ChartPlotInput.Line3D(
            false, ExcelChartGrouping.STANDARD, Optional.empty(), List.of(series)),
        new ChartPlotInput.Pie(false, Optional.empty(), List.of(series)),
        new ChartPlotInput.Pie3D(false, List.of(series)),
        new ChartPlotInput.Radar(false, ExcelChartRadarStyle.STANDARD, List.of(series)),
        new ChartPlotInput.Scatter(false, ExcelChartScatterStyle.LINE, List.of(series)),
        new ChartPlotInput.Surface(false, false, List.of(series)),
        new ChartPlotInput.Surface3D(false, false, List.of(series)));
  }

  private static ChartSeriesInput series(ChartTitleInput title) {
    return new ChartSeriesInput(
        title,
        new ChartDataSourceInput.StringLiteral(List.of("Jan", "Feb")),
        new ChartDataSourceInput.NumericLiteral(List.of(10d, 12d)),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static DataValidationInput validation() {
    return new DataValidationInput(
        new DataValidationRuleInput.WholeNumber(
            ExcelComparisonOperator.BETWEEN, "1", Optional.of("10")),
        false,
        false,
        Optional.of(
            new DataValidationPromptInput(
                TextSourceInput.standardInput(), TextSourceInput.inline("Prompt"), true)),
        Optional.of(
            new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceInput.inline("Alert"),
                TextSourceInput.standardInput(),
                true)));
  }

  private static PivotTableInput pivotTable() {
    return new PivotTableInput(
        "Sales Pivot 2026",
        new SheetName("Report"),
        new PivotTableInput.PivotSourceSpec.Range(new SheetName("Data"), new RangeRef("A1:D5")),
        new PivotTableInput.Anchor(new CellAddress("C5")),
        List.of("Region"),
        List.of("Stage"),
        List.of(),
        List.of(
            new PivotTableInput.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Amount", Optional.empty())));
  }

  private static TableInput table(TextSourceInput comment) {
    return new TableInput(
        new TableName("BudgetTable"),
        new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
        new RangeRef("A1:C4"),
        false,
        false,
        new TableStyleInput.None(),
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
