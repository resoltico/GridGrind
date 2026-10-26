package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsReport;
import dev.erst.gridgrind.core.model.layout.SheetDisplayReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.print.*;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookReadResult constructor invariants and defensive copies. */
class WorkbookReadResultTest {
  @Test
  void createsEachResultVariantAndCopiesCollections() {
    ReadResultFixture fixture = createReadResultFixture();
    fixture.clearSources();
    assertCopiedFixture(fixture);
  }

  @Test
  void validatesNestedReadResultInvariants() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSummary.WithSheets(
                -1,
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                0,
                true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSummary.WithSheets(
                1,
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                -1,
                true));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookSummary.WithSheets(
                1,
                null,
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                0,
                true));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCoreResult.NamedRangesResult(new StepId("ranges"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCoreResult.NamedRangesResult(
                new StepId("ranges"), List.of((NamedRangeReport) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetSummaryReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                ExcelSheetVisibility.VISIBLE,
                new SheetProtectionReport.Unprotected(),
                -1,
                0,
                0));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookSheetResult.CellsResult(
                new StepId("cells"), "Budget", null, WorkbookReadTestSupport.projection(), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.Window("Budget", "A1", 0, 1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.Window("Budget", "A1", 1, 0, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellHyperlinkReport(
                new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellCommentReport(new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookRuleResult.DataValidationsResult(
                new StepId("validations"), "Budget", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
                new StepId("conditionalFormattingHealth"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookRuleResult.AutofiltersResult(new StepId("autofilters"), "Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookRuleResult.TablesResult(new StepId("tables"), null));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(0, 0, 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetLayoutReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                new PaneReport.None(),
                9,
                defaultSheetPresentationSnapshot(),
                List.of(),
                List.of()));
    assertInstanceOf(PaneSpec.None.class, new PaneSpec.None());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSheetResult.PrintLayoutResult(new StepId("print"), "Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnLayoutReport(-1, 1.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnLayoutReport(0, 0.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnLayoutReport(0, Double.POSITIVE_INFINITY, false, 0, false));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnLayoutReport(0, 1.0, false, -1, false));
    assertThrows(
        IllegalArgumentException.class, () -> new RowLayoutReport(-1, 1.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class, () -> new RowLayoutReport(0, 0.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowLayoutReport(0, Double.POSITIVE_INFINITY, false, 0, false));
    assertThrows(
        IllegalArgumentException.class, () -> new RowLayoutReport(0, 1.0, false, -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.FormulaSurface(-1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetFormulaSurface("Budget", -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetFormulaSurface("Budget", 0, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.FormulaPattern("SUM(B2:B3)", 0, List.of("B4")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSurfaceResult.FormulaPattern("SUM(B2:B3)", 1, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 0, 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 1, 1, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                -1,
                "A1",
                "Item",
                0,
                0,
                List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                "TEXT"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                0,
                "A1",
                "Item",
                -1,
                0,
                List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                "TEXT"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                0,
                "A1",
                "Item",
                0,
                -1,
                List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                "TEXT"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSurfaceResult.TypeCount("TEXT", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(-1, 0, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, -1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, 0, -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, 0, 0, -1, List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, 0, 0, 0, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookSurfaceResult.NamedRangeSurfaceEntry(
                "BudgetTotal",
                null,
                "Budget!$B$4",
                WorkbookSurfaceResult.NamedRangeBackingKind.RANGE));
  }

  @Test
  void pivotReadResultsCopyEntriesAndRejectInvalidState() {
    PivotTableReport.Supported pivot =
        new PivotTableReport.Supported(
            "Budget Pivot",
            new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
            new PivotTableReport.Anchor(
                new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A3:C8")),
            new PivotTableReport.PivotSourceReport.Range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:D5")),
            List.of(new PivotTableReport.Field(0, "Region")),
            List.of(),
            List.of(),
            List.of(
                new PivotTableReport.DataField(
                    3,
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    Optional.of("#,##0.00"))),
            false);
    WorkbookAnalysis.AnalysisFinding finding =
        new WorkbookAnalysis.AnalysisFinding(
            AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
            AnalysisSeverity.WARNING,
            "Pivot table name is missing",
            "GridGrind assigned a synthetic identifier for readback.",
            new WorkbookAnalysis.AnalysisLocation.Sheet("Report"),
            List.of("_GG_PIVOT_Report_A3"));
    WorkbookAnalysis.PivotTableHealth health =
        new WorkbookAnalysis.PivotTableHealth(
            1, new WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0), List.of(finding));

    WorkbookDrawingResult.PivotTablesResult pivotTables =
        new WorkbookDrawingResult.PivotTablesResult(new StepId("pivots"), List.of(pivot));
    WorkbookAnalysisResult.PivotTableHealthResult pivotTableHealth =
        new WorkbookAnalysisResult.PivotTableHealthResult(new StepId("pivot-health"), health);

    assertEquals("Budget Pivot", pivotTables.pivotTables().getFirst().name());
    assertEquals(1, pivotTableHealth.analysis().checkedPivotTableCount());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingResult.PivotTablesResult(new StepId("pivots"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysisResult.PivotTableHealthResult(new StepId("pivot-health"), null));
  }

  @Test
  void reportsExcelNativeIndexDiagnosticsForReadResults() {
    CellReport.BlankReport blank =
        new CellReport.BlankReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            Optional.ofNullable(""),
            Optional.ofNullable(defaultStyle()),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Blank()));

    IllegalArgumentException windowRowFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookSheetResult.WindowRow(-1, List.of(blank)));
    assertTrue(windowRowFailure.getMessage().contains("rowIndex -1"));
    assertTrue(windowRowFailure.getMessage().contains("Excel row 1"));

    IllegalArgumentException schemaColumnFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookSurfaceResult.SchemaColumn(
                    -1,
                    "A1",
                    "Item",
                    0,
                    0,
                    List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                    "TEXT"));
    assertTrue(schemaColumnFailure.getMessage().contains("columnIndex -1"));
    assertTrue(schemaColumnFailure.getMessage().contains("Excel column A"));

    IllegalArgumentException lastRowFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new SheetSummaryReport(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                    ExcelSheetVisibility.VISIBLE,
                    new SheetProtectionReport.Unprotected(),
                    0,
                    -2,
                    0));
    assertTrue(lastRowFailure.getMessage().contains("lastRowIndex -2"));
    assertTrue(lastRowFailure.getMessage().contains("empty sheets report -1"));

    IllegalArgumentException lastColumnFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new SheetSummaryReport(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                    ExcelSheetVisibility.VISIBLE,
                    new SheetProtectionReport.Unprotected(),
                    0,
                    0,
                    -2));
    assertTrue(lastColumnFailure.getMessage().contains("lastColumnIndex -2"));
    assertTrue(lastColumnFailure.getMessage().contains("empty sheets report -1"));
  }

  private static CellStyleReport defaultStyle() {
    return new CellStyleReport(
        "",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            FontHeightReport.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None()),
        new CellProtectionReport(true, false));
  }

  private ReadResultFixture createReadResultFixture() {
    CellReport.BlankReport blank =
        new CellReport.BlankReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            Optional.ofNullable(""),
            Optional.ofNullable(defaultStyle()),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Blank()));
    List<String> sheetNames = new ArrayList<>(List.of("Budget"));
    List<NamedRangeReport> namedRanges =
        new ArrayList<>(
            List.of(
                new NamedRangeReport.RangeReport(
                    new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                    new NamedRangeScopeReport.Workbook(),
                    "Budget!$B$4",
                    NamedRangeTarget.range(
                        new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                        new dev.erst.gridgrind.core.vocabulary.RangeRef("B4")))));
    List<CellReport> cells = new ArrayList<>(List.of(blank));
    List<WorkbookSheetResult.WindowRow> rows =
        new ArrayList<>(List.of(new WorkbookSheetResult.WindowRow(0, List.of(blank))));
    List<MergedRegionReport> mergedRegions =
        new ArrayList<>(
            List.of(
                new MergedRegionReport(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"))));
    List<CellHyperlinkReport> hyperlinks =
        new ArrayList<>(
            List.of(
                new CellHyperlinkReport(
                    new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                    new HyperlinkTarget.Url("https://example.com/report"))));
    List<CellCommentReport> comments =
        new ArrayList<>(
            List.of(
                new CellCommentReport(
                    new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                    new CommentReport(
                        "Review", "GridGrind", false, Optional.empty(), Optional.empty()))));
    List<ColumnLayoutReport> columns =
        new ArrayList<>(List.of(new ColumnLayoutReport(0, 12.5, false, 0, false)));
    List<RowLayoutReport> resultRows =
        new ArrayList<>(List.of(new RowLayoutReport(0, 18.0, false, 0, false)));
    List<DataValidationEntryReport> validations =
        new ArrayList<>(
            List.of(
                new DataValidationEntryReport.Supported(
                    List.of("A2:A5").stream()
                        .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                        .toList(),
                    dev.erst.gridgrind.excel.DataValidationFixtures.report(
                        new ResolvedDataValidation(
                            new DataValidationRuleSpec.TextLength(
                                ExcelComparisonOperator.LESS_OR_EQUAL,
                                ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
                                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                                    Optional.empty(), "formula2")),
                            true,
                            false,
                            Optional.of(
                                new ResolvedValidationPrompt(
                                    "Reason", "Use 20 characters or fewer.", true)),
                            Optional.empty())))));
    List<AutofilterEntryReport> autofilters =
        new ArrayList<>(
            List.of(
                new AutofilterEntryReport.SheetOwned(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("E1:F4")),
                new AutofilterEntryReport.TableOwned(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:C4"), "BudgetTable")));
    List<TableEntryReport> tables =
        new ArrayList<>(
            List.of(
                new TableEntryReport(
                    new dev.erst.gridgrind.core.vocabulary.TableLabel("BudgetTable"),
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:C4"),
                    1,
                    1,
                    List.of("Item", "Amount", "Billable"),
                    new TableStyleReport.Named("TableStyleMedium2", false, false, true, false),
                    true)));
    List<WorkbookSurfaceResult.FormulaPattern> formulas =
        new ArrayList<>(
            List.of(new WorkbookSurfaceResult.FormulaPattern("SUM(B2:B3)", 1, List.of("B4"))));
    List<WorkbookSurfaceResult.SchemaColumn> schemaColumns =
        new ArrayList<>(
            List.of(
                new WorkbookSurfaceResult.SchemaColumn(
                    0,
                    "A1",
                    "Item",
                    2,
                    0,
                    List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 2)),
                    "TEXT")));
    List<WorkbookSurfaceResult.NamedRangeSurfaceEntry> namedRangeEntries =
        new ArrayList<>(
            List.of(
                new WorkbookSurfaceResult.NamedRangeSurfaceEntry(
                    "BudgetTotal",
                    new NamedRangeScopeReport.Workbook(),
                    "Budget!$B$4",
                    WorkbookSurfaceResult.NamedRangeBackingKind.RANGE)));

    return new ReadResultFixture(
        sheetNames,
        namedRanges,
        cells,
        rows,
        mergedRegions,
        hyperlinks,
        comments,
        columns,
        resultRows,
        validations,
        autofilters,
        tables,
        formulas,
        schemaColumns,
        namedRangeEntries,
        new WorkbookCoreResult.WorkbookSummaryResult(
            new StepId("workbook"),
            new WorkbookSummary.WithSheets(
                1,
                sheetNames.stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                1,
                true)),
        new WorkbookCoreResult.NamedRangesResult(new StepId("ranges"), namedRanges),
        new WorkbookSheetResult.SheetSummaryResult(
            new StepId("sheet"),
            new SheetSummaryReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                ExcelSheetVisibility.VISIBLE,
                new SheetProtectionReport.Unprotected(),
                4,
                3,
                2)),
        WorkbookReadTestSupport.cellsResult("cells", "Budget", cells),
        WorkbookReadTestSupport.windowResult(
            "window", new WorkbookSheetResult.Window("Budget", "A1", 1, 1, rows)),
        new WorkbookSheetResult.MergedRegionsResult(new StepId("merged"), "Budget", mergedRegions),
        new WorkbookSheetResult.HyperlinksResult(new StepId("hyperlinks"), "Budget", hyperlinks),
        new WorkbookSheetResult.CommentsResult(new StepId("comments"), "Budget", comments),
        new WorkbookSheetResult.SheetLayoutResult(
            new StepId("layout"),
            new SheetLayoutReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
                new PaneReport.Frozen(1, 1, 1, 1),
                125,
                defaultSheetPresentationSnapshot(),
                columns,
                resultRows)),
        new WorkbookSheetResult.PrintLayoutResult(
            new StepId("print"), "Budget", defaultPrintLayoutSnapshot()),
        new WorkbookRuleResult.DataValidationsResult(
            new StepId("validations"), "Budget", validations),
        new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
            new StepId("conditionalFormattingHealth"),
            new WorkbookAnalysis.ConditionalFormattingHealth(
                1, new WorkbookAnalysis.AnalysisSummary(0, 0, 0, 0), List.of())),
        new WorkbookRuleResult.AutofiltersResult(new StepId("autofilters"), "Budget", autofilters),
        new WorkbookRuleResult.TablesResult(new StepId("tables"), tables),
        new WorkbookSurfaceResult.FormulaSurfaceResult(
            new StepId("formula"),
            new WorkbookSurfaceResult.FormulaSurface(
                1,
                List.of(new WorkbookSurfaceResult.SheetFormulaSurface("Budget", 1, 1, formulas)))),
        WorkbookReadTestSupport.sheetSchemaResult(
            "schema",
            new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 3, 2, 2, schemaColumns)),
        new WorkbookSurfaceResult.NamedRangeSurfaceResult(
            new StepId("surface"),
            new WorkbookSurfaceResult.NamedRangeSurface(1, 0, 1, 0, namedRangeEntries)));
  }

  private void assertCopiedFixture(ReadResultFixture fixture) {
    assertEquals(
        List.of(new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
        fixture.workbookSummary().workbook().sheetNames());
    assertEquals(1, fixture.namedRangesResult().namedRanges().size());
    assertEquals(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
        fixture.sheetSummary().sheet().sheetName());
    assertEquals("A1", fixture.cellsResult().cells().getFirst().address().value());
    assertEquals(
        "A1",
        fixture.windowResult().window().rows().getFirst().cells().getFirst().address().value());
    assertEquals("A1:B2", fixture.mergedRegionsResult().mergedRegions().getFirst().range().value());
    assertEquals(
        "https://example.com/report",
        assertInstanceOf(
                HyperlinkTarget.Url.class,
                fixture.hyperlinksResult().hyperlinks().getFirst().hyperlink())
            .target());
    assertEquals("Review", fixture.commentsResult().comments().getFirst().comment().text());
    assertEquals(
        new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(1, 1, 1, 1),
        fixture.layoutResult().layout().pane());
    assertEquals(125, fixture.layoutResult().layout().zoomPercent());
    assertEquals(defaultPrintLayoutSnapshot(), fixture.printLayoutResult().printLayout());
    assertEquals(
        "A2:A5",
        fixture.dataValidationsResult().validations().getFirst().ranges().getFirst().value());
    assertEquals(
        "conditionalFormattingHealth",
        fixture.conditionalFormattingHealthResult().stepId().value());
    assertEquals("E1:F4", fixture.autofiltersResult().autofilters().getFirst().range().value());
    assertEquals("BudgetTable", fixture.tablesResult().tables().getFirst().name().value());
    assertEquals(1, fixture.formulaSurfaceResult().surface().totalFormulaCellCount());
    assertEquals("TEXT", fixture.sheetSchemaResult().surface().columns().getFirst().dominantType());
    assertEquals(
        WorkbookSurfaceResult.NamedRangeBackingKind.RANGE,
        fixture.namedRangeSurfaceResult().surface().namedRanges().getFirst().kind());
  }

  private record ReadResultFixture(
      List<String> sheetNames,
      List<NamedRangeReport> namedRanges,
      List<CellReport> cells,
      List<WorkbookSheetResult.WindowRow> rows,
      List<MergedRegionReport> mergedRegions,
      List<CellHyperlinkReport> hyperlinks,
      List<CellCommentReport> comments,
      List<ColumnLayoutReport> columns,
      List<RowLayoutReport> resultRows,
      List<DataValidationEntryReport> validations,
      List<AutofilterEntryReport> autofilters,
      List<TableEntryReport> tables,
      List<WorkbookSurfaceResult.FormulaPattern> formulas,
      List<WorkbookSurfaceResult.SchemaColumn> schemaColumns,
      List<WorkbookSurfaceResult.NamedRangeSurfaceEntry> namedRangeEntries,
      WorkbookCoreResult.WorkbookSummaryResult workbookSummary,
      WorkbookCoreResult.NamedRangesResult namedRangesResult,
      WorkbookSheetResult.SheetSummaryResult sheetSummary,
      WorkbookSheetResult.CellsResult cellsResult,
      WorkbookSheetResult.WindowResult windowResult,
      WorkbookSheetResult.MergedRegionsResult mergedRegionsResult,
      WorkbookSheetResult.HyperlinksResult hyperlinksResult,
      WorkbookSheetResult.CommentsResult commentsResult,
      WorkbookSheetResult.SheetLayoutResult layoutResult,
      WorkbookSheetResult.PrintLayoutResult printLayoutResult,
      WorkbookRuleResult.DataValidationsResult dataValidationsResult,
      WorkbookAnalysisResult.ConditionalFormattingHealthResult conditionalFormattingHealthResult,
      WorkbookRuleResult.AutofiltersResult autofiltersResult,
      WorkbookRuleResult.TablesResult tablesResult,
      WorkbookSurfaceResult.FormulaSurfaceResult formulaSurfaceResult,
      WorkbookSurfaceResult.SheetSchemaResult sheetSchemaResult,
      WorkbookSurfaceResult.NamedRangeSurfaceResult namedRangeSurfaceResult) {
    private void clearSources() {
      sheetNames.clear();
      namedRanges.clear();
      cells.clear();
      rows.clear();
      mergedRegions.clear();
      hyperlinks.clear();
      comments.clear();
      columns.clear();
      resultRows.clear();
      validations.clear();
      autofilters.clear();
      tables.clear();
      formulas.clear();
      schemaColumns.clear();
      namedRangeEntries.clear();
    }
  }

  private static PrintLayoutReport defaultPrintLayoutSnapshot() {
    return new PrintLayoutReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
        new PrintAreaReport.Range(new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B20")),
        ExcelPrintOrientation.LANDSCAPE,
        new PrintScalingReport.Fit(1, 0),
        new PrintTitleRowsReport.Band(0, 0),
        new PrintTitleColumnsReport.Band(0, 0),
        new HeaderFooterTextReport("Budget", "", "2026"),
        new HeaderFooterTextReport("", "Confidential", ""),
        new PrintSetupReport(
            new PrintMarginsReport(0.0d, 0.0d, 0.0d, 0.0d, 0.0d, 0.0d),
            false,
            false,
            false,
            0,
            false,
            false,
            0,
            false,
            0,
            List.of(),
            List.of()));
  }

  private static SheetPresentationReport defaultSheetPresentationSnapshot() {
    return new SheetPresentationReport(
        SheetDisplayReport.defaults(),
        Optional.empty(),
        SheetOutlineSummaryReport.defaults(),
        SheetDefaultsReport.defaults(),
        List.of());
  }
}
