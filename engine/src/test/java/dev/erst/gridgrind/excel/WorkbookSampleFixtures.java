package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.Base64;
import java.util.List;
import java.util.Optional;

/** Representative workbook command and read-command samples for exhaustive coverage tests. */
final class WorkbookSampleFixtures {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

  private WorkbookSampleFixtures() {}

  static List<WorkbookReadCommand.Introspection> introspectionCommands() {
    return List.of(
        new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook")),
        new WorkbookReadCommand.GetWorkbookProtection(new StepId("protection")),
        new WorkbookReadCommand.GetCustomXmlMappings(new StepId("customXmlMappings")),
        new WorkbookReadCommand.ExportCustomXmlMapping(
            new StepId("customXmlExport"),
            new CustomXmlMappingLocator(
                java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
            false,
            StandardCharsets.UTF_8.name()),
        new WorkbookReadCommand.GetNamedRanges(
            new StepId("ranges"), new ExcelNamedRangeSelection.All()),
        new WorkbookReadCommand.GetSheetSummary(new StepId("sheet"), "Budget"),
        WorkbookReadTestSupport.getCells("cells", "Budget", List.of("A1")),
        WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 1, 1),
        new WorkbookReadCommand.GetMergedRegions(new StepId("merged"), "Budget"),
        new WorkbookReadCommand.GetHyperlinks(
            new StepId("hyperlinks"), "Budget", new ExcelCellSelection.AllUsedCells()),
        new WorkbookReadCommand.GetComments(
            new StepId("comments"), "Budget", new ExcelCellSelection.Selected(List.of("A1"))),
        new WorkbookReadCommand.GetDrawingObjects(new StepId("drawings"), "Budget"),
        new WorkbookReadCommand.GetCharts(
            new StepId("charts"), new ExcelChartSelection.AllOnSheet("Budget")),
        new WorkbookReadCommand.GetPivotTables(
            new StepId("pivots"), new ExcelPivotTableSelection.All()),
        new WorkbookReadCommand.GetDrawingObjectPayload(
            new StepId("payload"), "Budget", "Object 1"),
        new WorkbookReadCommand.GetSheetLayout(new StepId("layout"), "Budget"),
        new WorkbookReadCommand.GetPrintLayout(new StepId("print"), "Budget"),
        new WorkbookReadCommand.GetDataValidations(
            new StepId("validations"), "Budget", new ExcelRangeSelection.All()),
        new WorkbookReadCommand.GetConditionalFormatting(
            new StepId("conditional"),
            "Budget",
            new ExcelRangeSelection.Selected(List.of("A1:B2"))),
        new WorkbookReadCommand.GetAutofilters(new StepId("autofilters"), "Budget"),
        new WorkbookReadCommand.GetTables(new StepId("tables"), new ExcelTableSelection.All()),
        new WorkbookReadCommand.GetArrayFormulas(
            new StepId("arrayFormulas"), new ExcelSheetSelection.Selected(List.of("Budget"))));
  }

  static List<WorkbookCommand> workbookCommands() {
    DrawingAnchorSpec.TwoCell firstAnchor = anchor(1, 1, 4, 6);
    DrawingAnchorSpec.TwoCell movedAnchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(6, 2, 0, 0),
            new DrawingMarkerSpec(9, 7, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);

    return List.of(
        new WorkbookSheetCommand.CreateSheet("Budget"),
        new WorkbookSheetCommand.RenameSheet("Budget", "Summary"),
        new WorkbookSheetCommand.DeleteSheet("Archive"),
        new WorkbookSheetCommand.MoveSheet("Budget", 1),
        new WorkbookSheetCommand.CopySheet(
            "Budget", "Budget Copy", new ExcelSheetCopyPosition.AppendAtEnd()),
        new WorkbookSheetCommand.SetActiveSheet("Budget"),
        new WorkbookSheetCommand.SetSelectedSheets(List.of("Budget", "Summary")),
        new WorkbookSheetCommand.SetSheetVisibility("Budget", ExcelSheetVisibility.VERY_HIDDEN),
        new WorkbookSheetCommand.SetSheetProtection("Budget", protectionSettings()),
        new WorkbookSheetCommand.ClearSheetProtection("Budget"),
        new WorkbookSheetCommand.SetWorkbookProtection(
            new ResolvedWorkbookProtection(
                true, false, true, Optional.of("book"), Optional.of("review"))),
        new WorkbookSheetCommand.ClearWorkbookProtection(),
        new WorkbookStructureCommand.MergeCells("Budget", "A1:B2"),
        new WorkbookStructureCommand.UnmergeCells("Budget", "A1:B2"),
        new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 1, 16.0d),
        new WorkbookStructureCommand.SetRowHeight("Budget", 0, 2, 28.5d),
        new WorkbookStructureCommand.InsertRows("Budget", 2, 3),
        new WorkbookStructureCommand.DeleteRows("Budget", new ExcelRowSpan(4, 6)),
        new WorkbookStructureCommand.ShiftRows("Budget", new ExcelRowSpan(1, 3), 2),
        new WorkbookStructureCommand.InsertColumns("Budget", 1, 2),
        new WorkbookStructureCommand.DeleteColumns("Budget", new ExcelColumnSpan(3, 4)),
        new WorkbookStructureCommand.ShiftColumns("Budget", new ExcelColumnSpan(0, 1), -1),
        new WorkbookStructureCommand.SetRowVisibility("Budget", new ExcelRowSpan(5, 7), true),
        new WorkbookStructureCommand.SetColumnVisibility(
            "Budget", new ExcelColumnSpan(2, 3), false),
        new WorkbookStructureCommand.GroupRows("Budget", new ExcelRowSpan(8, 10), false),
        new WorkbookStructureCommand.UngroupRows("Budget", new ExcelRowSpan(8, 10)),
        new WorkbookStructureCommand.GroupColumns("Budget", new ExcelColumnSpan(4, 6), true),
        new WorkbookStructureCommand.UngroupColumns("Budget", new ExcelColumnSpan(4, 6)),
        new WorkbookLayoutCommand.SetSheetPane("Budget", new ExcelSheetPane.Frozen(1, 2, 1, 2)),
        new WorkbookLayoutCommand.SetSheetZoom("Budget", 135),
        new WorkbookLayoutCommand.SetPrintLayout("Budget", defaultPrintLayout()),
        new WorkbookLayoutCommand.ClearPrintLayout("Budget"),
        new WorkbookCellCommand.SetCell(
            "Budget", "A1", new CellValueSpec.Date(LocalDate.of(2026, 3, 23))),
        new WorkbookCellCommand.SetArrayFormula(
            "Budget", "D2:D4", new ResolvedArrayFormula("B2:B4*C2:C4")),
        new WorkbookCellCommand.ClearArrayFormula("Budget", "D2"),
        new WorkbookMetadataCommand.ImportCustomXmlMapping(
            new ResolvedCustomXmlImport(
                new CustomXmlMappingLocator(
                    java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                "<CORSO><NOME>Ops</NOME></CORSO>")),
        new WorkbookCellCommand.SetRange(
            "Budget",
            "A1:B2",
            List.of(
                List.of(new CellValue.Text("Item"), new CellValue.NumberValue(49.0d)),
                List.of(new CellValue.Text("Tax"), new CellValue.NumberValue(10.0d)))),
        new WorkbookCellCommand.ClearRange("Budget", "C1:C2"),
        new WorkbookAnnotationCommand.SetHyperlink(
            "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")),
        new WorkbookAnnotationCommand.ClearHyperlink("Budget", "A1"),
        new WorkbookAnnotationCommand.SetComment(
            "Budget", "A1", new ResolvedComment("Review", "GridGrind", false)),
        new WorkbookAnnotationCommand.ClearComment("Budget", "A1"),
        new WorkbookDrawingCommand.SetPicture(
            "Budget",
            new ResolvedPicture(
                "BudgetPicture",
                new BinaryContent(PNG_PIXEL_BYTES),
                ExcelPictureFormat.PNG,
                firstAnchor,
                Optional.of("Queue preview"))),
        new WorkbookDrawingCommand.SetSignatureLine(
            "Budget",
            new ResolvedSignatureLine(
                "BudgetSignature",
                firstAnchor,
                false,
                Optional.ofNullable("Review the budget before signing."),
                Optional.ofNullable("Ada Lovelace"),
                Optional.ofNullable("Finance"),
                Optional.ofNullable("ada@example.com"),
                Optional.empty(),
                Optional.ofNullable("invalid"),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.of(new BinaryContent(PNG_PIXEL_BYTES)))),
        new WorkbookDrawingCommand.SetChart("Budget", chartDefinition(firstAnchor)),
        new WorkbookTabularCommand.SetPivotTable(
            new PivotTableSpec(
                "Budget Pivot",
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new PivotTableSpec.PivotSourceSpec.Range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:C4")),
                new PivotTableSpec.Anchor(new dev.erst.gridgrind.core.vocabulary.CellAddress("E3")),
                List.of("Item"),
                List.of(),
                List.of(),
                List.of(
                    new PivotTableSpec.DataField(
                        "Tax",
                        ExcelPivotDataConsolidateFunction.SUM,
                        "Total Tax",
                        Optional.of("#,##0.00"))))),
        new WorkbookDrawingCommand.SetShape(
            "Budget",
            new ResolvedShape.SimpleShape(
                "BudgetShape", firstAnchor, "rect", Optional.of("Queue"))),
        new WorkbookDrawingCommand.SetEmbeddedObject(
            "Budget",
            new ResolvedEmbeddedObject(
                "BudgetEmbed",
                "Payload",
                "payload.txt",
                "payload.txt",
                new BinaryContent("payload".getBytes(StandardCharsets.UTF_8)),
                ExcelPictureFormat.PNG,
                new BinaryContent(PNG_PIXEL_BYTES),
                firstAnchor)),
        new WorkbookDrawingCommand.SetDrawingObjectAnchor("Budget", "BudgetShape", movedAnchor),
        new WorkbookDrawingCommand.DeleteDrawingObject("Budget", "BudgetPicture"),
        new WorkbookFormattingCommand.ApplyStyle("Budget", "A1:B1", defaultStyle()),
        new WorkbookFormattingCommand.SetDataValidation("Budget", "B2:B5", validationDefinition()),
        new WorkbookFormattingCommand.ClearDataValidations(
            "Budget", new ExcelRangeSelection.Selected(List.of("C2:D4"))),
        new WorkbookFormattingCommand.SetConditionalFormatting(
            "Budget",
            new ExcelConditionalFormattingBlockDefinition(
                List.of("A2:A5"),
                List.of(
                    new ConditionalFormattingRuleSpec.FormulaRule(
                        "A2>0",
                        true,
                        Optional.of(
                            new DifferentialStylePatch(
                                Optional.of("0.00"),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty())))))),
        new WorkbookFormattingCommand.ClearConditionalFormatting(
            "Budget", new ExcelRangeSelection.Selected(List.of("A2:A5"))),
        new WorkbookTabularCommand.SetAutofilter(
            "Budget",
            "A1:C4",
            List.of(
                new AutofilterFilterColumnSpec(
                    0L, false, new AutofilterFilterCriterionSpec.Values(List.of("Ada"), true))),
            Optional.of(
                new AutofilterSortStateSpec(
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:C4"),
                    true,
                    false,
                    Optional.empty(),
                    List.of(
                        new AutofilterSortConditionSpec.Value(
                            new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A4"), false))))),
        new WorkbookTabularCommand.ClearAutofilter("Budget"),
        new WorkbookTabularCommand.SetTable(
            new ResolvedTable(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef("A1:C4"),
                true,
                new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false))),
        new WorkbookTabularCommand.DeleteTable("BudgetTable", "Budget"),
        new WorkbookTabularCommand.DeletePivotTable("Budget Pivot", "Budget"),
        new WorkbookMetadataCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Budget", "B4"))),
        new WorkbookMetadataCommand.DeleteNamedRange(
            "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
        new WorkbookCellCommand.AppendRow("Budget", List.of(new CellValue.Text("Item"))),
        new WorkbookLayoutCommand.AutoSizeColumns("Budget"));
  }

  static DrawingAnchorSpec.TwoCell anchor(int fromRow, int fromColumn, int toRow, int toColumn) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(fromRow, fromColumn, 0, 0),
        new DrawingMarkerSpec(toRow, toColumn, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static ResolvedChart chartDefinition(DrawingAnchorSpec.TwoCell anchor) {
    return ExcelChartTestSupport.barChart(
        "BudgetChart",
        anchor,
        new ResolvedChartTitle.Text("Roadmap"),
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ResolvedChart.Series(
                new ChartTitleSpec.Formula("B1"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"))));
  }

  private static ResolvedDataValidation validationDefinition() {
    return new ResolvedDataValidation(
        new DataValidationRuleSpec.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL,
            ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
            ExcelComparisonFormulaSupport.normalizeOptionalFormula(Optional.empty(), "formula2")),
        true,
        false,
        Optional.of(new ResolvedValidationPrompt("Reason", "Use 20 characters or fewer.", true)),
        Optional.of(
            new ResolvedValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP, "Too long", "Use a shorter reason.", true)));
  }

  private static CellStylePatch defaultStyle() {
    return new CellStylePatch(
        Optional.of("#,##0.00"),
        Optional.of(
            new CellAlignmentPatch(
                Optional.of(true),
                Optional.of(ExcelHorizontalAlignment.RIGHT),
                Optional.of(ExcelVerticalAlignment.CENTER),
                Optional.empty(),
                Optional.empty())),
        Optional.of(
            new CellFontPatch(
                Optional.of(true),
                Optional.of(false),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static ResolvedPrintLayout defaultPrintLayout() {
    return new ResolvedPrintLayout(
        new PrintAreaSpec.Range(new RangeRef("A1:C20")),
        ExcelPrintOrientation.LANDSCAPE,
        new PrintScalingSpec.Fit(1, 0),
        new PrintTitleRowsSpec.Band(0, 1),
        new PrintTitleColumnsSpec.Band(0, 0),
        new HeaderFooterTextReport("Left", "Center", "Right"),
        new HeaderFooterTextReport("Footer Left", "", "Footer Right"));
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }
}
