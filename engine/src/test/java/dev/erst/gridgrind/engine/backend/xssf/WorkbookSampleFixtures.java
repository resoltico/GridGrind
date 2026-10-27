package dev.erst.gridgrind.engine.backend.xssf;

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
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
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
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.ResolvedTarget;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.formulas.ComparisonFormulas;
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

  static List<dev.erst.gridgrind.core.plan.ResolvedInspectionStep> introspectionCommands() {
    return List.of(
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("workbook"),
            new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
            new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetWorkbookSummary()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("protection"),
            new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
            new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetWorkbookProtection()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("customXmlMappings"),
            new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
            new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetCustomXmlMappings()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("customXmlExport"),
            new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
            new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                new CustomXmlMappingLocator(
                    java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                false,
                StandardCharsets.UTF_8.name())),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("ranges"),
            new dev.erst.gridgrind.core.plan.NamedRangeSelector.All(),
            new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetNamedRanges()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("sheet"),
            new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetSummary()),
        WorkbookReadTestSupport.getCells("cells", "Budget", List.of("A1")),
        WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 1, 1),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("merged"),
            new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetMergedRegions()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("hyperlinks"),
            new dev.erst.gridgrind.core.plan.CellSelector.AllUsedInSheet(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetHyperlinks()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("comments"),
            new dev.erst.gridgrind.core.plan.CellSelector.ByAddresses(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                List.of("A1").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
                    .toList()),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetComments()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("drawings"),
            new dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("charts"),
            new ChartSelector.AllOnSheet(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetCharts()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("pivots"),
            new PivotTableSelector.All(),
            new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetPivotTables()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("payload"),
            new dev.erst.gridgrind.core.plan.DrawingObjectSelector.ByName(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"), "Object 1"),
            new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                .GetDrawingObjectPayload()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("layout"),
            new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetLayout()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("print"),
            new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetPrintLayout()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("validations"),
            new dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetDataValidations()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("conditional"),
            new dev.erst.gridgrind.core.plan.RangeSelector.ByRanges(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                List.of("A1:B2").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                    .toList()),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetConditionalFormatting()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("autofilters"),
            new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetAutofilters()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("tables"),
            new TableSelector.All(),
            new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetTables()),
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("arrayFormulas"),
            new dev.erst.gridgrind.core.plan.SheetSelector.ByNames(
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetName::new)
                    .toList()),
            new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetArrayFormulas()));
  }

  static List<ResolvedStep> workbookCommands() {
    DrawingAnchorSpec.TwoCell firstAnchor = anchor(1, 1, 4, 6);
    DrawingAnchorSpec.TwoCell movedAnchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(6, 2, 0, 0),
            new DrawingMarkerSpec(9, 7, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);

    return List.of(
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.EnsureSheet()),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.RenameSheet(new SheetName("Summary"))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Archive")),
            new WorkbookMutationAction.DeleteSheet()),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.MoveSheet(1)),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.CopySheet(
                new SheetName("Budget Copy"), new SheetCopyPosition.AppendAtEnd())),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.SetActiveSheet()),
        new ResolvedStep(
            new SheetSelector.ByNames(
                List.of("Budget", "Summary").stream().map(SheetName::new).toList()),
            new WorkbookMutationAction.SetSelectedSheets()),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.SetSheetVisibility(ExcelSheetVisibility.VERY_HIDDEN)),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.SetSheetProtection(protectionSettings(), Optional.empty())),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.ClearSheetProtection()),
        new ResolvedStep(
            new WorkbookSelector.Current(),
            new ResolvedMutationAction.SetWorkbookProtection(
                new ResolvedWorkbookProtection(
                    true, false, true, Optional.of("book"), Optional.of("review")))),
        new ResolvedStep(
            new WorkbookSelector.Current(), new WorkbookMutationAction.ClearWorkbookProtection()),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
            new WorkbookMutationAction.MergeCells()),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
            new WorkbookMutationAction.UnmergeCells()),
        new ResolvedStep(
            new ColumnBandSelector.Span(new SheetName("Budget"), 0, 1),
            new WorkbookMutationAction.SetColumnWidth(16.0d)),
        new ResolvedStep(
            new RowBandSelector.Span(new SheetName("Budget"), 0, 2),
            new WorkbookMutationAction.SetRowHeight(28.5d)),
        new ResolvedStep(
            new RowBandSelector.Insertion(new SheetName("Budget"), 2, 3),
            new WorkbookMutationAction.InsertRows()),
        new ResolvedStep(
            new RowBandSelector.Span(new SheetName("Budget"), 4, 6),
            new WorkbookMutationAction.DeleteRows()),
        new ResolvedStep(
            new RowBandSelector.Span(new SheetName("Budget"), 1, 3),
            new WorkbookMutationAction.ShiftRows(2)),
        new ResolvedStep(
            new ColumnBandSelector.Insertion(new SheetName("Budget"), 1, 2),
            new WorkbookMutationAction.InsertColumns()),
        new ResolvedStep(
            new ColumnBandSelector.Span(new SheetName("Budget"), 3, 4),
            new WorkbookMutationAction.DeleteColumns()),
        new ResolvedStep(
            new ColumnBandSelector.Span(new SheetName("Budget"), 0, 1),
            new WorkbookMutationAction.ShiftColumns(-1)),
        new ResolvedStep(
            new RowBandSelector.Span(new SheetName("Budget"), 5, 7),
            new WorkbookMutationAction.SetRowVisibility(true)),
        new ResolvedStep(
            new ColumnBandSelector.Span(new SheetName("Budget"), 2, 3),
            new WorkbookMutationAction.SetColumnVisibility(false)),
        new ResolvedStep(
            new RowBandSelector.Span(new SheetName("Budget"), 8, 10),
            new WorkbookMutationAction.GroupRows(false)),
        new ResolvedStep(
            new RowBandSelector.Span(new SheetName("Budget"), 8, 10),
            new WorkbookMutationAction.UngroupRows()),
        new ResolvedStep(
            new ColumnBandSelector.Span(new SheetName("Budget"), 4, 6),
            new WorkbookMutationAction.GroupColumns(true)),
        new ResolvedStep(
            new ColumnBandSelector.Span(new SheetName("Budget"), 4, 6),
            new WorkbookMutationAction.UngroupColumns()),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.SetSheetPane(new PaneSpec.Frozen(1, 2, 1, 2))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.SetSheetZoom(135)),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.SetPrintLayout(defaultPrintLayout())),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.ClearPrintLayout()),
        new ResolvedStep(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new ResolvedMutationAction.SetCell(new CellValueSpec.Date(LocalDate.of(2026, 3, 23)))),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("D2:D4")),
            new ResolvedMutationAction.SetArrayFormula(new ResolvedArrayFormula("B2:B4*C2:C4"))),
        new ResolvedStep(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("D2")),
            new CellMutationAction.ClearArrayFormula()),
        new ResolvedStep(
            new WorkbookSelector.Current(),
            new ResolvedMutationAction.ImportCustomXmlMapping(
                new ResolvedCustomXmlImport(
                    new CustomXmlMappingLocator(
                        java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                    "<CORSO><NOME>Ops</NOME></CORSO>"))),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
            new ResolvedMutationAction.SetRange(
                List.of(
                    List.of(new CellValue.Text("Item"), new CellValue.NumberValue(49.0d)),
                    List.of(new CellValue.Text("Tax"), new CellValue.NumberValue(10.0d))))),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("C1:C2")),
            new CellMutationAction.ClearRange()),
        new ResolvedStep(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.SetHyperlink(
                new HyperlinkTarget.Url("https://example.com/report"))),
        new ResolvedStep(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.ClearHyperlink()),
        new ResolvedStep(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new ResolvedMutationAction.SetComment(
                new ResolvedComment("Review", "GridGrind", false))),
        new ResolvedStep(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.ClearComment()),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.SetPicture(
                new ResolvedPicture(
                    "BudgetPicture",
                    new BinaryContent(PNG_PIXEL_BYTES),
                    ExcelPictureFormat.PNG,
                    firstAnchor,
                    Optional.of("Queue preview")))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.SetSignatureLine(
                new ResolvedSignatureLine(
                    "BudgetSignature",
                    firstAnchor,
                    false,
                    Optional.of("Review the budget before signing."),
                    Optional.of("Ada Lovelace"),
                    Optional.of("Finance"),
                    Optional.of("ada@example.com"),
                    Optional.empty(),
                    Optional.of("invalid"),
                    Optional.of(ExcelPictureFormat.PNG),
                    Optional.of(new BinaryContent(PNG_PIXEL_BYTES))))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.SetChart(chartDefinition(firstAnchor))),
        setPivotTableStep(
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
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.SetShape(
                new ResolvedShape.SimpleShape(
                    "BudgetShape", firstAnchor, "rect", Optional.of("Queue")))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.SetEmbeddedObject(
                new ResolvedEmbeddedObject(
                    "BudgetEmbed",
                    "Payload",
                    "payload.txt",
                    "payload.txt",
                    new BinaryContent("payload".getBytes(StandardCharsets.UTF_8)),
                    ExcelPictureFormat.PNG,
                    new BinaryContent(PNG_PIXEL_BYTES),
                    firstAnchor))),
        new ResolvedStep(
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "BudgetShape"),
            new DrawingMutationAction.SetDrawingObjectAnchor(movedAnchor)),
        new ResolvedStep(
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "BudgetPicture"),
            new DrawingMutationAction.DeleteDrawingObject()),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
            new CellMutationAction.ApplyStyle(defaultStyle())),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("B2:B5")),
            new ResolvedMutationAction.SetDataValidation(validationDefinition())),
        new dev.erst.gridgrind.core.plan.ResolvedStep(
            new dev.erst.gridgrind.core.plan.RangeSelector.ByRanges(
                new SheetName("Budget"),
                List.of("C2:D4").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                    .toList()),
            new StructuredMutationAction.ClearDataValidations()),
        new ResolvedStep(
            new RangeSelector.ByRanges(
                new SheetName("Budget"),
                List.of("A2:A5").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                    .toList()),
            new StructuredMutationAction.SetConditionalFormatting(
                new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
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
                                    Optional.empty()))))))),
        new dev.erst.gridgrind.core.plan.ResolvedStep(
            new dev.erst.gridgrind.core.plan.RangeSelector.ByRanges(
                new SheetName("Budget"),
                List.of("A2:A5").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                    .toList()),
            new StructuredMutationAction.ClearConditionalFormatting()),
        new ResolvedStep(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:C4")),
            new StructuredMutationAction.SetAutofilter(
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
                                new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A4"),
                                false)))))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new StructuredMutationAction.ClearAutofilter()),
        setTableStep(
            new ResolvedTable(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef("A1:C4"),
                true,
                new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false))),
        new ResolvedStep(
            new TableSelector.ByNameOnSheet(new TableName("BudgetTable"), new SheetName("Budget")),
            new StructuredMutationAction.DeleteTable()),
        new ResolvedStep(
            new PivotTableSelector.ByNameOnSheet("Budget Pivot", new SheetName("Budget")),
            new StructuredMutationAction.DeletePivotTable()),
        setNamedRangeStep(
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B4")))),
        deleteNamedRangeStep(
            "BudgetTotal",
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new ResolvedMutationAction.AppendRow(List.of(new CellValue.Text("Item")))),
        new ResolvedStep(
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.AutoSizeColumns()));
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
            ComparisonFormulas.normalizeFormula("20", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")),
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

  static ResolvedStep setTableStep(dev.erst.gridgrind.core.model.table.ResolvedTable table) {
    return new ResolvedStep(
        new TableSelector.ByNameOnSheet(table.name(), table.sheetName()),
        new ResolvedMutationAction.SetTable(table));
  }

  static ResolvedStep setPivotTableStep(dev.erst.gridgrind.core.model.pivot.PivotTableSpec pivot) {
    return new ResolvedStep(
        new PivotTableSelector.ByNameOnSheet(pivot.name(), pivot.sheetName()),
        new StructuredMutationAction.SetPivotTable(pivot));
  }

  static ResolvedStep setNamedRangeStep(StructuredMutationAction.SetNamedRange action) {
    ResolvedTarget target =
        switch (action.scope()) {
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook _ ->
              new NamedRangeSelector.WorkbookScope(action.name());
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet scope ->
              new NamedRangeSelector.SheetScope(action.name(), scope.sheetName());
        };
    return new ResolvedStep(target, action);
  }

  static ResolvedStep deleteNamedRangeStep(
      String name, dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport scope) {
    ResolvedTarget target =
        switch (scope) {
          case dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Workbook _ ->
              new NamedRangeSelector.WorkbookScope(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName(name));
          case dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Sheet sheet ->
              new NamedRangeSelector.SheetScope(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName(name),
                  new SheetName(sheet.sheetName().value()));
        };
    return new ResolvedStep(target, new StructuredMutationAction.DeleteNamedRange());
  }
}
