package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.print.*;
import dev.erst.gridgrind.core.model.print.PrintAreaReport;
import dev.erst.gridgrind.core.model.print.PrintScalingReport;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.NamedRangeBackingKind;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Workbook read-result translation coverage. */
class DefaultGridGrindRequestExecutorReadResultTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void convertsReadResultsIntoProtocolReadResults() {
    CellReport.BlankReport blank =
        new CellReport.BlankReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            Optional.ofNullable(""),
            Optional.ofNullable(defaultStyle()),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Blank()));

    InspectionResult workbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult(
                new StepId("workbook"),
                new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.WithSheets(
                    1, List.of("Budget"), "Budget", List.of("Budget"), 1, true)));
    InspectionResult namedRanges =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookCoreResult.NamedRangesResult(
                new StepId("ranges"),
                List.of(
                    new NamedRangeReport.RangeReport(
                        new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                        new NamedRangeScopeReport.Workbook(),
                        "Budget!$B$4",
                        NamedRangeTarget.range(
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                            new dev.erst.gridgrind.core.vocabulary.RangeRef("B4"))))));
    InspectionResult sheetSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult(
                new StepId("sheet"),
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility.VISIBLE,
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Unprotected(),
                    4,
                    3,
                    2)));
    InspectionResult cells =
        InspectionResultConverter.toReadResult(excelCellsResult("cells", "Budget", List.of(blank)));
    InspectionResult window =
        InspectionResultConverter.toReadResult(
            excelWindowResult(
                "window",
                new dev.erst.gridgrind.excel.WorkbookSheetResult.Window(
                    "Budget",
                    "A1",
                    1,
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSheetResult.WindowRow(
                            0, List.of(blank))))));
    InspectionResult merged =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.MergedRegionsResult(
                new StepId("merged"),
                "Budget",
                List.of(new dev.erst.gridgrind.excel.WorkbookSheetResult.MergedRegion("A1:B2"))));
    InspectionResult hyperlinks =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.HyperlinksResult(
                new StepId("hyperlinks"),
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.CellHyperlink(
                        "A1", new HyperlinkTarget.Url("https://example.com/report")))));
    InspectionResult comments =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.CommentsResult(
                new StepId("comments"),
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.CellComment(
                        "A1",
                        new CommentReport(
                            "Review",
                            "GridGrind",
                            false,
                            java.util.Optional.empty(),
                            java.util.Optional.empty())))));
    InspectionResult layout =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                new StepId("layout"),
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout(
                    "Budget",
                    new dev.erst.gridgrind.core.model.layout.PaneSpec.Frozen(1, 1, 1, 1),
                    125,
                    defaultSheetPresentationSnapshot(),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSheetResult.ColumnLayout(
                            0, 12.5, false, 0, false)),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSheetResult.RowLayout(
                            0, 18.0, false, 0, false)))));
    InspectionResult printLayout =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.PrintLayoutResult(
                new StepId("printLayout"),
                "Budget",
                new dev.erst.gridgrind.core.model.print.PrintLayoutReport(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                    new dev.erst.gridgrind.core.model.print.PrintAreaReport.Range(
                        new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B20")),
                    ExcelPrintOrientation.LANDSCAPE,
                    new dev.erst.gridgrind.core.model.print.PrintScalingReport.Fit(1, 0),
                    new dev.erst.gridgrind.core.model.print.PrintTitleRowsReport.Band(0, 0),
                    new dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport.Band(0, 0),
                    new dev.erst.gridgrind.core.model.print.HeaderFooterTextReport(
                        "Budget", "", ""),
                    new dev.erst.gridgrind.core.model.print.HeaderFooterTextReport(
                        "", "Page &P", ""),
                    defaultPrintSetupSnapshot())));
    InspectionResult conditionalFormatting =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult(
                new StepId("conditional-formatting"),
                "Budget",
                List.of(
                    new ConditionalFormattingEntryReport(
                        List.of("A2:A5").stream()
                            .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                            .toList(),
                        List.of(
                            new ConditionalFormattingRuleReport.FormulaRule(
                                1,
                                true,
                                "A2>0",
                                Optional.ofNullable(
                                    new DifferentialStyleReport(
                                        "0.00",
                                        true,
                                        null,
                                        null,
                                        Optional.ofNullable(
                                            dev.erst.gridgrind.excel.DifferentialStyleFixtures
                                                .color(ColorSpec.rgb(new RgbHex("#102030")))),
                                        null,
                                        null,
                                        Optional.ofNullable(
                                            dev.erst.gridgrind.excel.DifferentialStyleFixtures
                                                .color(ColorSpec.rgb(new RgbHex("#E0F0AA")))),
                                        Optional.empty(),
                                        List.of()))),
                            new ConditionalFormattingRuleReport.ColorScaleRule(
                                2,
                                false,
                                List.of(
                                    new ConditionalFormattingThresholdReport(
                                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                                    new ConditionalFormattingThresholdReport(
                                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                                List.of("#AA0000", "#00AA00").stream()
                                    .map(dev.erst.gridgrind.core.vocabulary.RgbHex::new)
                                    .toList()))))));
    InspectionResult formulaSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaSurfaceResult(
                new StepId("formula"),
                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaSurface(
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetFormulaSurface(
                            "Budget",
                            1,
                            1,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaPattern(
                                    "SUM(B2:B3)", 1, List.of("B4"))))))));
    InspectionResult schema =
        InspectionResultConverter.toReadResult(
            excelSheetSchemaResult(
                "schema",
                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchema(
                    "Budget",
                    "A1",
                    3,
                    2,
                    2,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSurfaceResult.SchemaColumn(
                            0,
                            "A1",
                            "Item",
                            2,
                            0,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.TypeCount(
                                    "TEXT", 2)),
                            "TEXT")))));
    InspectionResult namedRangeSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceResult(
                new StepId("surface"),
                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurface(
                    1,
                    0,
                    1,
                    0,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceEntry(
                            "BudgetTotal",
                            new NamedRangeScopeReport.Workbook(),
                            "Budget!$B$4",
                            dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeBackingKind
                                .RANGE)))));
    InspectionResult formulaHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.FormulaHealthResult(
                new StepId("formula-health"),
                new dev.erst.gridgrind.excel.WorkbookAnalysis.FormulaHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                                .FORMULA_VOLATILE_FUNCTION,
                            dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.INFO,
                            "Volatile formula",
                            "Formula uses NOW().",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Cell(
                                "Budget", "B4"),
                            List.of("NOW()"))))));

    assertInstanceOf(WorkbookInspectionResult.WorkbookSummaryResult.class, workbookSummary);
    assertInstanceOf(WorkbookInspectionResult.NamedRangesResult.class, namedRanges);
    assertInstanceOf(SheetInspectionResult.SheetSummaryResult.class, sheetSummary);
    assertInstanceOf(SheetInspectionResult.CellsResult.class, cells);
    assertInstanceOf(SheetInspectionResult.WindowResult.class, window);
    assertInstanceOf(SheetInspectionResult.MergedRegionsResult.class, merged);
    assertInstanceOf(SheetInspectionResult.HyperlinksResult.class, hyperlinks);
    assertInstanceOf(SheetInspectionResult.CommentsResult.class, comments);
    assertInstanceOf(SheetInspectionResult.SheetLayoutResult.class, layout);
    assertInstanceOf(SheetInspectionResult.PrintLayoutResult.class, printLayout);
    assertInstanceOf(
        SheetInspectionResult.ConditionalFormattingResult.class, conditionalFormatting);
    assertInstanceOf(WorkbookSurfaceInspectionResult.FormulaSurfaceResult.class, formulaSurface);
    assertInstanceOf(WorkbookSurfaceInspectionResult.SheetSchemaResult.class, schema);
    assertInstanceOf(
        WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface);
    assertInstanceOf(
        dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult.class,
        formulaHealth);
    assertEquals(
        "Budget",
        cast(WorkbookInspectionResult.WorkbookSummaryResult.class, workbookSummary)
            .workbook()
            .sheetNames()
            .getFirst()
            .value());
    assertEquals(
        "BudgetTotal",
        cast(WorkbookInspectionResult.NamedRangesResult.class, namedRanges)
            .namedRanges()
            .getFirst()
            .name()
            .value());
    assertEquals(
        "Budget",
        cast(SheetInspectionResult.SheetSummaryResult.class, sheetSummary)
            .sheet()
            .sheetName()
            .value());
    assertEquals(
        "A1",
        cast(SheetInspectionResult.CellsResult.class, cells).cells().getFirst().address().value());
    assertEquals(
        "A1",
        windowCells(cast(SheetInspectionResult.WindowResult.class, window).window())
            .getFirst()
            .address()
            .value());
    assertEquals(
        "A1:B2",
        cast(SheetInspectionResult.MergedRegionsResult.class, merged)
            .mergedRegions()
            .getFirst()
            .range()
            .value());
    assertEquals(
        new HyperlinkTarget.Url("https://example.com/report"),
        cast(SheetInspectionResult.HyperlinksResult.class, hyperlinks)
            .hyperlinks()
            .getFirst()
            .hyperlink());
    assertEquals(
        "Review",
        cast(SheetInspectionResult.CommentsResult.class, comments)
            .comments()
            .getFirst()
            .comment()
            .text());
    assertInstanceOf(
        PaneReport.Frozen.class,
        cast(SheetInspectionResult.SheetLayoutResult.class, layout).layout().pane());
    assertEquals(
        125, cast(SheetInspectionResult.SheetLayoutResult.class, layout).layout().zoomPercent());
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        cast(SheetInspectionResult.PrintLayoutResult.class, printLayout).layout().orientation());
    assertEquals(
        2,
        cast(SheetInspectionResult.ConditionalFormattingResult.class, conditionalFormatting)
            .conditionalFormattingBlocks()
            .getFirst()
            .rules()
            .size());
    assertEquals(
        1,
        cast(WorkbookSurfaceInspectionResult.FormulaSurfaceResult.class, formulaSurface)
            .surface()
            .totalFormulaCellCount());
    assertEquals(
        "TEXT",
        cast(WorkbookSurfaceInspectionResult.SheetSchemaResult.class, schema)
            .surface()
            .columns()
            .getFirst()
            .dominantType());
    assertEquals(
        NamedRangeBackingKind.RANGE,
        cast(WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface)
            .surface()
            .namedRanges()
            .getFirst()
            .kind());
    assertEquals(
        1,
        cast(
                dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult.class,
                formulaHealth)
            .analysis()
            .summary()
            .infoCount());
  }

  @Test
  void convertsRemainingAnalysisReadResultsIntoProtocolReadResults() {
    InspectionResult conditionalFormattingHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.ConditionalFormattingHealthResult(
                new StepId("conditional-formatting-health"),
                new dev.erst.gridgrind.excel.WorkbookAnalysis.ConditionalFormattingHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 1, 0, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                                .CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
                            dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
                            "Priority collision",
                            "Conditional-formatting priorities collide.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("FORMULA_RULE@Budget!A1:A3"))))));
    InspectionResult hyperlinkHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.HyperlinkHealthResult(
                new StepId("hyperlink-health"),
                new dev.erst.gridgrind.excel.WorkbookAnalysis.HyperlinkHealth(
                    2,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(2, 1, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                                .HYPERLINK_MALFORMED_TARGET,
                            dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
                            "Malformed hyperlink target",
                            "Hyperlink target is malformed.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .Workbook(),
                            List.of("mailto:")),
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                                .HYPERLINK_MISSING_DOCUMENT_SHEET,
                            dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.WARNING,
                            "Missing target sheet",
                            "Sheet is missing.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("Missing!A1"))))));
    InspectionResult namedRangeHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.NamedRangeHealthResult(
                new StepId("named-range-health"),
                new dev.erst.gridgrind.excel.WorkbookAnalysis.NamedRangeHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                                .NAMED_RANGE_BROKEN_REFERENCE,
                            dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.WARNING,
                            "Broken named range",
                            "Named range contains #REF!.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Range(
                                "Budget", "A1:B2"),
                            List.of("#REF!"))))));
    InspectionResult workbookFindings =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.WorkbookFindingsResult(
                new StepId("workbook-findings"),
                new dev.erst.gridgrind.excel.WorkbookAnalysis.WorkbookFindings(
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                                .NAMED_RANGE_SCOPE_SHADOWING,
                            dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.INFO,
                            "Scope shadowing",
                            "Name exists in more than one scope.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .NamedRange(
                                "BudgetTotal",
                                new NamedRangeScopeReport.Sheet(
                                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"))),
                            List.of("Budget!$B$4"))))));

    assertEquals(
        1,
        cast(
                dev.erst.gridgrind.core.plan.WorkbookAnalysisResult
                    .ConditionalFormattingHealthResult.class,
                conditionalFormattingHealth)
            .analysis()
            .checkedConditionalFormattingBlockCount());
    AnalysisLocationReport workbookLocation =
        cast(
                dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.HyperlinkHealthResult.class,
                hyperlinkHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    AnalysisLocationReport sheetLocation =
        cast(
                dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.HyperlinkHealthResult.class,
                hyperlinkHealth)
            .analysis()
            .findings()
            .get(1)
            .location();
    AnalysisLocationReport rangeLocation =
        cast(
                dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.NamedRangeHealthResult.class,
                namedRangeHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    AnalysisLocationReport namedRangeLocation =
        cast(
                dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.WorkbookFindingsResult.class,
                workbookFindings)
            .analysis()
            .findings()
            .getFirst()
            .location();

    assertInstanceOf(AnalysisLocationReport.Workbook.class, workbookLocation);
    assertEquals(
        "Budget", cast(AnalysisLocationReport.Sheet.class, sheetLocation).sheetName().value());
    assertEquals("A1:B2", cast(AnalysisLocationReport.Range.class, rangeLocation).range());
    assertEquals(
        "BudgetTotal", cast(AnalysisLocationReport.NamedRange.class, namedRangeLocation).name());
  }

  @Test
  void convertsNamedRangeFormulaSnapshotsAndFormulaBackedSurfaceEntries() {
    WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult surface =
        cast(
            WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceResult(
                    new StepId("surface"),
                    new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurface(
                        0,
                        1,
                        0,
                        1,
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookSurfaceResult
                                .NamedRangeSurfaceEntry(
                                "BudgetRollup",
                                new NamedRangeScopeReport.Sheet(
                                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
                                "SUM(Budget!$B$2:$B$3)",
                                dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeBackingKind
                                    .FORMULA))))));

    assertEquals(NamedRangeBackingKind.FORMULA, surface.surface().namedRanges().getFirst().kind());
  }

  @Test
  void convertsPaneNoneIntoProtocolReport() {
    SheetInspectionResult.SheetLayoutResult layout =
        cast(
            SheetInspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                    new StepId("layout"),
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.core.model.layout.PaneSpec.None(),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));

    assertInstanceOf(PaneReport.None.class, layout.layout().pane());
  }

  @Test
  void convertsSplitPaneAndDefaultPrintLayoutIntoProtocolReports() {
    SheetInspectionResult.SheetLayoutResult layout =
        cast(
            SheetInspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                    new StepId("layout"),
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.core.model.layout.PaneSpec.Split(
                            1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));
    SheetInspectionResult.PrintLayoutResult printLayout =
        cast(
            SheetInspectionResult.PrintLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.PrintLayoutResult(
                    new StepId("print-layout"),
                    "Budget",
                    new dev.erst.gridgrind.core.model.print.PrintLayoutReport(
                        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                        new dev.erst.gridgrind.core.model.print.PrintAreaReport.None(),
                        ExcelPrintOrientation.PORTRAIT,
                        new dev.erst.gridgrind.core.model.print.PrintScalingReport.Automatic(),
                        new dev.erst.gridgrind.core.model.print.PrintTitleRowsReport.None(),
                        new dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport.None(),
                        new dev.erst.gridgrind.core.model.print.HeaderFooterTextReport("", "", ""),
                        new dev.erst.gridgrind.core.model.print.HeaderFooterTextReport("", "", ""),
                        defaultPrintSetupSnapshot()))));

    PaneReport.Split pane = assertInstanceOf(PaneReport.Split.class, layout.layout().pane());
    assertEquals(1200, pane.xSplitPosition());
    assertEquals(2400, pane.ySplitPosition());
    assertEquals(3, pane.leftmostColumn());
    assertEquals(4, pane.topRow());
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, pane.activePane());
    assertInstanceOf(PrintAreaReport.None.class, printLayout.layout().printArea());
    assertInstanceOf(PrintScalingReport.Automatic.class, printLayout.layout().scaling());
    assertInstanceOf(PrintTitleRowsReport.None.class, printLayout.layout().repeatingRows());
    assertInstanceOf(PrintTitleColumnsReport.None.class, printLayout.layout().repeatingColumns());
  }

  @Test
  void convertsSheetStateReadResultsIntoProtocolShapes() {
    InspectionResult emptyWorkbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult(
                new StepId("workbook"),
                new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.Empty(
                    0, List.of(), 0, false)));
    InspectionResult populatedWorkbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult(
                new StepId("workbook-2"),
                new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.WithSheets(
                    2,
                    List.of("Budget", "Archive"),
                    "Archive",
                    List.of("Budget", "Archive"),
                    1,
                    true)));
    InspectionResult protectedSheetSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult(
                new StepId("sheet"),
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility.VERY_HIDDEN,
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Protected(
                        excelProtectionSettings()),
                    4,
                    7,
                    3)));

    WorkbookSummary.Empty empty =
        assertInstanceOf(
            WorkbookSummary.Empty.class,
            cast(WorkbookInspectionResult.WorkbookSummaryResult.class, emptyWorkbookSummary)
                .workbook());
    WorkbookSummary.WithSheets populated =
        assertInstanceOf(
            WorkbookSummary.WithSheets.class,
            cast(WorkbookInspectionResult.WorkbookSummaryResult.class, populatedWorkbookSummary)
                .workbook());
    SheetSummaryReport sheet =
        cast(SheetInspectionResult.SheetSummaryResult.class, protectedSheetSummary).sheet();
    SheetProtectionReport.Protected protection =
        assertInstanceOf(SheetProtectionReport.Protected.class, sheet.protection());

    assertEquals(0, empty.sheetCount());
    assertEquals("Archive", populated.activeSheetName().value());
    assertEquals(
        List.of("Budget", "Archive"),
        populated.selectedSheetNames().stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
    assertEquals(ExcelSheetVisibility.VERY_HIDDEN, sheet.visibility());
    assertEquals(protectionSettings(), protection.settings());
  }
}
