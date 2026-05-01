package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports;
import dev.erst.gridgrind.contract.dto.GridGrindSchemaAndFormulaReports;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Workbook read-result translation coverage. */
class DefaultGridGrindRequestExecutorReadResultTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void convertsReadResultsIntoProtocolReadResults() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    InspectionResult workbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.WithSheets(
                    1, List.of("Budget"), "Budget", List.of("Budget"), 1, true)));
    InspectionResult namedRanges =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookCoreResult.NamedRangesResult(
                "ranges",
                List.of(
                    new ExcelNamedRangeSnapshot.RangeSnapshot(
                        "BudgetTotal",
                        new ExcelNamedRangeScope.WorkbookScope(),
                        "Budget!$B$4",
                        new ExcelNamedRangeTarget("Budget", "B4")))));
    InspectionResult sheetSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VISIBLE,
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Unprotected(),
                    4,
                    3,
                    2)));
    InspectionResult cells =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.CellsResult(
                "cells", "Budget", List.of(blank)));
    InspectionResult window =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.WindowResult(
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
                "merged",
                "Budget",
                List.of(new dev.erst.gridgrind.excel.WorkbookSheetResult.MergedRegion("A1:B2"))));
    InspectionResult hyperlinks =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.HyperlinksResult(
                "hyperlinks",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.CellHyperlink(
                        "A1", new ExcelHyperlink.Url("https://example.com/report")))));
    InspectionResult comments =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.CommentsResult(
                "comments",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.CellComment(
                        "A1",
                        new ExcelCommentSnapshot("Review", "GridGrind", false, null, null)))));
    InspectionResult layout =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                "layout",
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout(
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelSheetPane.Frozen(1, 1, 1, 1),
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
                "printLayout",
                "Budget",
                new dev.erst.gridgrind.excel.ExcelPrintLayoutSnapshot(
                    new dev.erst.gridgrind.excel.ExcelPrintLayout(
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.Range("A1:B20"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Fit(1, 0),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.Band(0, 0),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.Band(0, 0),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("Budget", "", ""),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "Page &P", "")),
                    defaultPrintSetupSnapshot())));
    InspectionResult conditionalFormatting =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookRuleResult.ConditionalFormattingResult(
                "conditional-formatting",
                "Budget",
                List.of(
                    new ExcelConditionalFormattingBlockSnapshot(
                        List.of("A2:A5"),
                        List.of(
                            new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                                1,
                                true,
                                "A2>0",
                                new ExcelDifferentialStyleSnapshot(
                                    "0.00", true, null, null, "#102030", null, null, "#E0F0AA",
                                    null, List.of())),
                            new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
                                2,
                                false,
                                List.of(
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                                List.of("#AA0000", "#00AA00")))))));
    InspectionResult formulaSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSurfaceResult.FormulaSurfaceResult(
                "formula",
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
            new dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchemaResult(
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
                                    "STRING", 2)),
                            "STRING")))));
    InspectionResult namedRangeSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceResult(
                "surface",
                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurface(
                    1,
                    0,
                    1,
                    0,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceEntry(
                            "BudgetTotal",
                            new ExcelNamedRangeScope.WorkbookScope(),
                            "Budget!$B$4",
                            dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeBackingKind
                                .RANGE)))));
    InspectionResult formulaHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.FormulaHealthResult(
                "formula-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.FormulaHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .FORMULA_VOLATILE_FUNCTION,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.INFO,
                            "Volatile formula",
                            "Formula uses NOW().",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Cell(
                                "Budget", "B4"),
                            List.of("NOW()"))))));

    assertInstanceOf(InspectionResult.WorkbookSummaryResult.class, workbookSummary);
    assertInstanceOf(InspectionResult.NamedRangesResult.class, namedRanges);
    assertInstanceOf(InspectionResult.SheetSummaryResult.class, sheetSummary);
    assertInstanceOf(InspectionResult.CellsResult.class, cells);
    assertInstanceOf(InspectionResult.WindowResult.class, window);
    assertInstanceOf(InspectionResult.MergedRegionsResult.class, merged);
    assertInstanceOf(InspectionResult.HyperlinksResult.class, hyperlinks);
    assertInstanceOf(InspectionResult.CommentsResult.class, comments);
    assertInstanceOf(InspectionResult.SheetLayoutResult.class, layout);
    assertInstanceOf(InspectionResult.PrintLayoutResult.class, printLayout);
    assertInstanceOf(InspectionResult.ConditionalFormattingResult.class, conditionalFormatting);
    assertInstanceOf(InspectionResult.FormulaSurfaceResult.class, formulaSurface);
    assertInstanceOf(InspectionResult.SheetSchemaResult.class, schema);
    assertInstanceOf(InspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface);
    assertInstanceOf(InspectionResult.FormulaHealthResult.class, formulaHealth);
    assertEquals(
        "Budget",
        cast(InspectionResult.WorkbookSummaryResult.class, workbookSummary)
            .workbook()
            .sheetNames()
            .getFirst());
    assertEquals(
        "BudgetTotal",
        cast(InspectionResult.NamedRangesResult.class, namedRanges)
            .namedRanges()
            .getFirst()
            .name());
    assertEquals(
        "Budget",
        cast(InspectionResult.SheetSummaryResult.class, sheetSummary).sheet().sheetName());
    assertEquals(
        "A1", cast(InspectionResult.CellsResult.class, cells).cells().getFirst().address());
    assertEquals(
        "A1",
        cast(InspectionResult.WindowResult.class, window)
            .window()
            .rows()
            .getFirst()
            .cells()
            .getFirst()
            .address());
    assertEquals(
        "A1:B2",
        cast(InspectionResult.MergedRegionsResult.class, merged)
            .mergedRegions()
            .getFirst()
            .range());
    assertEquals(
        new HyperlinkTarget.Url("https://example.com/report"),
        cast(InspectionResult.HyperlinksResult.class, hyperlinks)
            .hyperlinks()
            .getFirst()
            .hyperlink());
    assertEquals(
        "Review",
        cast(InspectionResult.CommentsResult.class, comments)
            .comments()
            .getFirst()
            .comment()
            .text());
    assertInstanceOf(
        PaneReport.Frozen.class,
        cast(InspectionResult.SheetLayoutResult.class, layout).layout().pane());
    assertEquals(
        125, cast(InspectionResult.SheetLayoutResult.class, layout).layout().zoomPercent());
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        cast(InspectionResult.PrintLayoutResult.class, printLayout).layout().orientation());
    assertEquals(
        2,
        cast(InspectionResult.ConditionalFormattingResult.class, conditionalFormatting)
            .conditionalFormattingBlocks()
            .getFirst()
            .rules()
            .size());
    assertEquals(
        1,
        cast(InspectionResult.FormulaSurfaceResult.class, formulaSurface)
            .analysis()
            .totalFormulaCellCount());
    assertEquals(
        "STRING",
        cast(InspectionResult.SheetSchemaResult.class, schema)
            .analysis()
            .columns()
            .getFirst()
            .dominantType());
    assertEquals(
        GridGrindSchemaAndFormulaReports.NamedRangeBackingKind.RANGE,
        cast(InspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface)
            .analysis()
            .namedRanges()
            .getFirst()
            .kind());
    assertEquals(
        1,
        cast(InspectionResult.FormulaHealthResult.class, formulaHealth)
            .analysis()
            .summary()
            .infoCount());
  }

  @Test
  void convertsRemainingAnalysisReadResultsIntoProtocolReadResults() {
    InspectionResult conditionalFormattingHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.ConditionalFormattingHealthResult(
                "conditional-formatting-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.ConditionalFormattingHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 1, 0, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
                            "Priority collision",
                            "Conditional-formatting priorities collide.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("FORMULA_RULE@Budget!A1:A3"))))));
    InspectionResult hyperlinkHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.HyperlinkHealthResult(
                "hyperlink-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.HyperlinkHealth(
                    2,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(2, 1, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .HYPERLINK_MALFORMED_TARGET,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
                            "Malformed hyperlink target",
                            "Hyperlink target is malformed.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .Workbook(),
                            List.of("mailto:")),
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .HYPERLINK_MISSING_DOCUMENT_SHEET,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.WARNING,
                            "Missing target sheet",
                            "Sheet is missing.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("Missing!A1"))))));
    InspectionResult namedRangeHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.NamedRangeHealthResult(
                "named-range-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.NamedRangeHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .NAMED_RANGE_BROKEN_REFERENCE,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.WARNING,
                            "Broken named range",
                            "Named range contains #REF!.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Range(
                                "Budget", "A1:B2"),
                            List.of("#REF!"))))));
    InspectionResult workbookFindings =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookAnalysisResult.WorkbookFindingsResult(
                "workbook-findings",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.WorkbookFindings(
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .NAMED_RANGE_SCOPE_SHADOWING,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.INFO,
                            "Scope shadowing",
                            "Name exists in more than one scope.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .NamedRange(
                                "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
                            List.of("Budget!$B$4"))))));

    assertEquals(
        1,
        cast(InspectionResult.ConditionalFormattingHealthResult.class, conditionalFormattingHealth)
            .analysis()
            .checkedConditionalFormattingBlockCount());
    GridGrindAnalysisReports.AnalysisLocationReport workbookLocation =
        cast(InspectionResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindAnalysisReports.AnalysisLocationReport sheetLocation =
        cast(InspectionResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .get(1)
            .location();
    GridGrindAnalysisReports.AnalysisLocationReport rangeLocation =
        cast(InspectionResult.NamedRangeHealthResult.class, namedRangeHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindAnalysisReports.AnalysisLocationReport namedRangeLocation =
        cast(InspectionResult.WorkbookFindingsResult.class, workbookFindings)
            .analysis()
            .findings()
            .getFirst()
            .location();

    assertInstanceOf(
        GridGrindAnalysisReports.AnalysisLocationReport.Workbook.class, workbookLocation);
    assertEquals(
        "Budget",
        cast(GridGrindAnalysisReports.AnalysisLocationReport.Sheet.class, sheetLocation)
            .sheetName());
    assertEquals(
        "A1:B2",
        cast(GridGrindAnalysisReports.AnalysisLocationReport.Range.class, rangeLocation).range());
    assertEquals(
        "BudgetTotal",
        cast(GridGrindAnalysisReports.AnalysisLocationReport.NamedRange.class, namedRangeLocation)
            .name());
  }

  @Test
  void convertsNamedRangeFormulaSnapshotsAndFormulaBackedSurfaceEntries() {
    GridGrindWorkbookSurfaceReports.NamedRangeReport formulaReport =
        InspectionResultWorkbookCoreReportSupport.toNamedRangeReport(
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "BudgetRollup", new ExcelNamedRangeScope.WorkbookScope(), "SUM(Budget!$B$2:$B$3)"));
    assertInstanceOf(
        GridGrindWorkbookSurfaceReports.NamedRangeReport.FormulaReport.class, formulaReport);

    InspectionResult.NamedRangeSurfaceResult surface =
        cast(
            InspectionResult.NamedRangeSurfaceResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurfaceResult(
                    "surface",
                    new dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeSurface(
                        0,
                        1,
                        0,
                        1,
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookSurfaceResult
                                .NamedRangeSurfaceEntry(
                                "BudgetRollup",
                                new ExcelNamedRangeScope.SheetScope("Budget"),
                                "SUM(Budget!$B$2:$B$3)",
                                dev.erst.gridgrind.excel.WorkbookSurfaceResult.NamedRangeBackingKind
                                    .FORMULA))))));

    assertEquals(
        GridGrindSchemaAndFormulaReports.NamedRangeBackingKind.FORMULA,
        surface.analysis().namedRanges().getFirst().kind());
  }

  @Test
  void convertsPaneNoneIntoProtocolReport() {
    InspectionResult.SheetLayoutResult layout =
        cast(
            InspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.None(),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));

    assertInstanceOf(PaneReport.None.class, layout.layout().pane());
  }

  @Test
  void convertsSplitPaneAndDefaultPrintLayoutIntoProtocolReports() {
    InspectionResult.SheetLayoutResult layout =
        cast(
            InspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.Split(
                            1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));
    InspectionResult.PrintLayoutResult printLayout =
        cast(
            InspectionResult.PrintLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookSheetResult.PrintLayoutResult(
                    "print-layout",
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelPrintLayoutSnapshot(
                        new dev.erst.gridgrind.excel.ExcelPrintLayout(
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None(),
                            ExcelPrintOrientation.PORTRAIT,
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic(),
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None(),
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None(),
                            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", ""),
                            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", "")),
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
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.Empty(
                    0, List.of(), 0, false)));
    InspectionResult populatedWorkbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult(
                "workbook-2",
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
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VERY_HIDDEN,
                    new dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Protected(
                        excelProtectionSettings()),
                    4,
                    7,
                    3)));

    GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty empty =
        assertInstanceOf(
            GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty.class,
            cast(InspectionResult.WorkbookSummaryResult.class, emptyWorkbookSummary).workbook());
    GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets populated =
        assertInstanceOf(
            GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets.class,
            cast(InspectionResult.WorkbookSummaryResult.class, populatedWorkbookSummary)
                .workbook());
    GridGrindWorkbookSurfaceReports.SheetSummaryReport sheet =
        cast(InspectionResult.SheetSummaryResult.class, protectedSheetSummary).sheet();
    GridGrindWorkbookSurfaceReports.SheetProtectionReport.Protected protection =
        assertInstanceOf(
            GridGrindWorkbookSurfaceReports.SheetProtectionReport.Protected.class,
            sheet.protection());

    assertEquals(0, empty.sheetCount());
    assertEquals("Archive", populated.activeSheetName());
    assertEquals(List.of("Budget", "Archive"), populated.selectedSheetNames());
    assertEquals(ExcelSheetVisibility.VERY_HIDDEN, sheet.visibility());
    assertEquals(protectionSettings(), protection.settings());
  }
}
