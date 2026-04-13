package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for GridGrindResponse record construction, defaults, and typed accessors. */
class GridGrindResponseTest {
  @Test
  void createsSuccessAndFailureResponses() {
    List<WorkbookReadResult> reads = new ArrayList<>();
    reads.add(
        new WorkbookReadResult.WorkbookSummaryResult(
            "workbook",
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), 1, true)));

    GridGrindResponse.Success success =
        new GridGrindResponse.Success(
            null,
            new GridGrindResponse.PersistenceOutcome.SavedAs("budget.xlsx", "/tmp/budget.xlsx"),
            List.of(new RequestWarning(0, "SET_CELL", "Quote spaced sheet names in formulas.")),
            reads);
    reads.clear();

    assertInstanceOf(GridGrindResponse.Success.class, success);
    assertEquals(GridGrindProtocolVersion.V1, success.protocolVersion());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.SavedAs.class, success.persistence());
    assertEquals(1, success.warnings().size());
    assertEquals(1, success.reads().size());

    GridGrindResponse.Failure failure =
        new GridGrindResponse.Failure(
            null,
            GridGrindResponse.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "Bad request",
                new GridGrindResponse.ProblemContext.ValidateRequest(null, null)));

    assertEquals(GridGrindProtocolVersion.V1, failure.protocolVersion());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals(List.of(), failure.problem().causes());
  }

  @Test
  void copiesAndValidatesNestedCollections() {
    List<String> sheetNames = new ArrayList<>(List.of("Budget"));
    GridGrindResponse.WorkbookSummary workbook =
        new GridGrindResponse.WorkbookSummary.WithSheets(
            1, sheetNames, "Budget", List.of("Budget"), 1, true);
    List<WorkbookReadResult> reads =
        new ArrayList<>(
            List.of(new WorkbookReadResult.WorkbookSummaryResult("workbook", workbook)));

    GridGrindResponse.Success success = new GridGrindResponse.Success(null, null, null, reads);
    sheetNames.clear();
    reads.clear();

    assertEquals(List.of("Budget"), workbook.sheetNames());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, success.persistence());
    assertEquals(List.of(), success.warnings());
    assertEquals(1, success.reads().size());

    assertThrows(
        NullPointerException.class,
        () -> new GridGrindResponse.Success(null, null, List.of(), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, null, "Budget", List.of("Budget"), 0, true));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.CellReport.TextReport(
                "A1", "STRING", "x", null, null, null, "x", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.NamedRangeReport.RangeReport(
                "BudgetTotal", new NamedRangeScope.Workbook(), "Budget!$B$4", null));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindResponse.PersistenceOutcome.SavedAs(null, "/tmp/budget.xlsx"));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindResponse.PersistenceOutcome.SavedAs("budget.xlsx", null));
    assertThrows(
        IllegalArgumentException.class, () -> new RequestWarning(-1, "SET_CELL", "warning"));
    assertThrows(IllegalArgumentException.class, () -> new RequestWarning(0, " ", "warning"));
    assertThrows(IllegalArgumentException.class, () -> new RequestWarning(0, "SET_CELL", " "));
  }

  @Test
  void workbookSummaryAndSheetProtectionVariantsExposeB1State() {
    GridGrindResponse.WorkbookSummary.Empty empty =
        new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), 0, false);
    SheetProtectionSettings protectionSettings = protectionSettings();
    GridGrindResponse.WorkbookSummary.WithSheets withSheets =
        new GridGrindResponse.WorkbookSummary.WithSheets(
            2, List.of("Budget", "Archive"), "Archive", List.of("Budget", "Archive"), 1, true);
    GridGrindResponse.SheetProtectionReport.Unprotected unprotected =
        new GridGrindResponse.SheetProtectionReport.Unprotected();
    GridGrindResponse.SheetProtectionReport.Protected protectedReport =
        new GridGrindResponse.SheetProtectionReport.Protected(protectionSettings);
    GridGrindResponse.SheetSummaryReport sheetSummary =
        new GridGrindResponse.SheetSummaryReport(
            "Budget", ExcelSheetVisibility.VERY_HIDDEN, protectedReport, 4, 8, 3);

    assertEquals(0, empty.sheetCount());
    assertEquals(List.of(), empty.sheetNames());
    assertEquals("Archive", withSheets.activeSheetName());
    assertEquals(List.of("Budget", "Archive"), withSheets.selectedSheetNames());
    assertNull(unprotected.settings());
    assertEquals(protectionSettings, protectedReport.settings());
    assertEquals(ExcelSheetVisibility.VERY_HIDDEN, sheetSummary.visibility());
    assertEquals(protectedReport, sheetSummary.protection());
  }

  @Test
  void problemContextRecordsReturnCorrectStages() {
    GridGrindResponse.ProblemContext.ParseArguments parseArguments =
        new GridGrindResponse.ProblemContext.ParseArguments("--response");
    assertEquals("PARSE_ARGUMENTS", parseArguments.stage());
    assertEquals("--response", parseArguments.argument());

    GridGrindResponse.ProblemContext.ReadRequest readRequest =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", "reads[0]", 7, 21);
    assertEquals("READ_REQUEST", readRequest.stage());
    assertEquals("/tmp/request.json", readRequest.requestPath());

    GridGrindResponse.ProblemContext.ApplyOperation applyOperation =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            "NEW", "SAVE_AS", 1, "SET_NAMED_RANGE", "Budget", "B4", "B4", null, "BudgetTotal");
    assertEquals("APPLY_OPERATION", applyOperation.stage());
    assertEquals("BudgetTotal", applyOperation.namedRangeName());

    GridGrindResponse.ProblemContext.ExecuteRead executeRead =
        new GridGrindResponse.ProblemContext.ExecuteRead(
            "NEW", "SAVE_AS", 2, "GET_CELLS", "cells", "Budget", "B4", "SUM(B2:B3)", "BudgetTotal");
    assertEquals("EXECUTE_READ", executeRead.stage());
    assertEquals("cells", executeRead.requestId());
    assertEquals("BudgetTotal", executeRead.namedRangeName());
  }

  @Test
  void withExceptionDataPreservesExistingProblemContextFields() {
    GridGrindResponse.ProblemContext.ApplyOperation applyBase =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            "NEW", "NONE", 0, "SET_CELL", null, null, null, null, null);
    GridGrindResponse.ProblemContext.ApplyOperation enrichedApply =
        applyBase.withExceptionData("Budget", "B4", null, "SUM(B2:B3)", "BudgetTotal");

    assertEquals("Budget", enrichedApply.sheetName());
    assertEquals("B4", enrichedApply.address());
    assertEquals("SUM(B2:B3)", enrichedApply.formula());
    assertEquals("BudgetTotal", enrichedApply.namedRangeName());

    GridGrindResponse.ProblemContext.ExecuteRead readBase =
        new GridGrindResponse.ProblemContext.ExecuteRead(
            "NEW", "NONE", 0, "GET_CELLS", "cells", null, null, null, null);
    GridGrindResponse.ProblemContext.ExecuteRead enrichedRead =
        readBase.withExceptionData("Budget", "C1", "SUM(B1:B3)", "BudgetTotal");

    assertEquals("Budget", enrichedRead.sheetName());
    assertEquals("C1", enrichedRead.address());
    assertEquals("SUM(B1:B3)", enrichedRead.formula());
    assertEquals("BudgetTotal", enrichedRead.namedRangeName());

    GridGrindResponse.ProblemContext.ApplyOperation preservedApply =
        new GridGrindResponse.ProblemContext.ApplyOperation(
                "NEW", "NONE", 0, "SET_CELL", "Budget", "A1", "A1:B2", "SUM(A1:A2)", "BudgetTotal")
            .withExceptionData("Other", "C3", "C3:D4", "NOW()", "OtherRange");
    assertEquals("Budget", preservedApply.sheetName());
    assertEquals("A1", preservedApply.address());
    assertEquals("A1:B2", preservedApply.range());
    assertEquals("SUM(A1:A2)", preservedApply.formula());
    assertEquals("BudgetTotal", preservedApply.namedRangeName());

    GridGrindResponse.ProblemContext.ExecuteRead preservedRead =
        new GridGrindResponse.ProblemContext.ExecuteRead(
                "NEW", "NONE", 0, "GET_CELLS", "cells", "Budget", "A1", "SUM(A1:A2)", "BudgetTotal")
            .withExceptionData("Other", "C3", "NOW()", "OtherRange");
    assertEquals("Budget", preservedRead.sheetName());
    assertEquals("A1", preservedRead.address());
    assertEquals("SUM(A1:A2)", preservedRead.formula());
    assertEquals("BudgetTotal", preservedRead.namedRangeName());
  }

  @Test
  void readRequestWithJsonPopulatesAndPreservesJsonLocationFields() {
    GridGrindResponse.ProblemContext.ReadRequest blank =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);
    GridGrindResponse.ProblemContext.ReadRequest enriched = blank.withJson("$.reads", 4, 12);

    assertEquals("$.reads", enriched.jsonPath());
    assertEquals(4, enriched.jsonLine());
    assertEquals(12, enriched.jsonColumn());

    GridGrindResponse.ProblemContext.ReadRequest preserved =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", "$.source", 1, 2)
            .withJson("$.reads", 4, 12);
    assertEquals("$.source", preserved.jsonPath());
    assertEquals(1, preserved.jsonLine());
    assertEquals(2, preserved.jsonColumn());
  }

  @Test
  void cellReportAndReadPayloadSubtypesExposeTypedGettersAndMetadataDefaults() {
    GridGrindResponse.CellStyleReport style = defaultCellStyleReport();
    HyperlinkTarget hyperlink = new HyperlinkTarget.Url("https://example.com");
    GridGrindResponse.CommentReport comment =
        new GridGrindResponse.CommentReport("Review", "GridGrind", false);

    GridGrindResponse.CellReport.BlankReport blankReport =
        new GridGrindResponse.CellReport.BlankReport("A1", "BLANK", "", style, null, null);
    GridGrindResponse.CellReport.FormulaReport formulaReport =
        new GridGrindResponse.CellReport.FormulaReport(
            "E1",
            "FORMULA",
            "85",
            style,
            hyperlink,
            comment,
            "SUM(B1:B2)",
            new GridGrindResponse.CellReport.NumberReport(
                "E1", "NUMBER", "85", style, null, null, 85.0));

    GridGrindResponse.WindowReport window =
        new GridGrindResponse.WindowReport(
            "Budget",
            "A1",
            1,
            1,
            List.of(new GridGrindResponse.WindowRowReport(0, List.of(blankReport))));
    GridGrindResponse.SheetLayoutReport layout =
        new GridGrindResponse.SheetLayoutReport(
            "Budget",
            new PaneReport.Frozen(1, 1, 1, 1),
            125,
            SheetPresentationReport.defaults(),
            List.of(new GridGrindResponse.ColumnLayoutReport(0, 12.5d, false, 0, false)),
            List.of(new GridGrindResponse.RowLayoutReport(0, 18.0d, false, 0, false)));
    PrintLayoutReport printLayout =
        new PrintLayoutReport(
            "Budget",
            new PrintAreaReport.Range("A1:B20"),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingReport.Fit(1, 0),
            new PrintTitleRowsReport.Band(0, 0),
            new PrintTitleColumnsReport.Band(0, 0),
            new HeaderFooterTextReport("Budget", "", ""),
            new HeaderFooterTextReport("", "Page &P", ""));

    assertEquals("FORMULA", formulaReport.effectiveType());
    assertEquals(
        "STRING",
        new GridGrindResponse.CellReport.TextReport(
                "B1", "STRING", "x", style, null, null, "x", null)
            .effectiveType());
    assertEquals(
        "NUMBER",
        new GridGrindResponse.CellReport.NumberReport("B2", "NUMBER", "12", style, null, null, 12.0)
            .effectiveType());
    assertEquals(
        "BOOLEAN",
        new GridGrindResponse.CellReport.BooleanReport(
                "B3", "BOOLEAN", "TRUE", style, null, null, true)
            .effectiveType());
    assertEquals(new HyperlinkTarget.Url("https://example.com"), formulaReport.hyperlink());
    assertEquals("Review", formulaReport.comment().text());
    assertNull(blankReport.formula());
    assertNull(blankReport.hyperlink());
    assertNull(blankReport.comment());
    assertNull(blankReport.stringValue());
    assertNull(blankReport.richText());
    assertNull(blankReport.numberValue());
    assertNull(blankReport.booleanValue());
    assertNull(blankReport.errorValue());
    assertEquals("A1", window.rows().getFirst().cells().getFirst().address());
    assertInstanceOf(PaneReport.Frozen.class, layout.pane());
    assertEquals(125, layout.zoomPercent());
    assertEquals(ExcelPrintOrientation.LANDSCAPE, printLayout.orientation());
  }

  @Test
  void textReportValidatesStructuredRichTextFacts() {
    GridGrindResponse.CellStyleReport style = defaultCellStyleReport();
    List<RichTextRunReport> richText =
        List.of(
            new RichTextRunReport(
                "Budget",
                new CellFontReport(
                    false,
                    false,
                    "Aptos",
                    new FontHeightReport(220, new BigDecimal("11")),
                    null,
                    false,
                    false)),
            new RichTextRunReport(
                " FY26",
                new CellFontReport(
                    true,
                    false,
                    "Aptos",
                    new FontHeightReport(220, new BigDecimal("11")),
                    rgb("#AABBCC"),
                    false,
                    false)));

    GridGrindResponse.CellReport.TextReport report =
        new GridGrindResponse.CellReport.TextReport(
            "A1", "STRING", "Budget FY26", style, null, null, "Budget FY26", richText);

    assertEquals(richText, report.richText());
    assertThrows(
        NullPointerException.class,
        () ->
            new RichTextRunReport(
                null,
                new CellFontReport(
                    false,
                    false,
                    "Aptos",
                    new FontHeightReport(220, new BigDecimal("11")),
                    null,
                    false,
                    false)));
    assertThrows(NullPointerException.class, () -> new RichTextRunReport("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RichTextRunReport(
                "",
                new CellFontReport(
                    false,
                    false,
                    "Aptos",
                    new FontHeightReport(220, new BigDecimal("11")),
                    null,
                    false,
                    false)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CellReport.TextReport(
                "A1", "STRING", "", style, null, null, "", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CellReport.TextReport(
                "A1", "STRING", "Mismatch", style, null, null, "Mismatch", richText));
  }

  @Test
  void namedRangeAndAnalysisReportsExposeExpectedDefaults() {
    GridGrindResponse.NamedRangeReport.FormulaReport formulaReport =
        new GridGrindResponse.NamedRangeReport.FormulaReport(
            "BudgetRollup", new NamedRangeScope.Workbook(), "SUM(Budget!$B$2:$B$3)");
    GridGrindResponse.NamedRangeSurfaceReport surface =
        new GridGrindResponse.NamedRangeSurfaceReport(
            1,
            0,
            0,
            1,
            List.of(
                new GridGrindResponse.NamedRangeSurfaceEntryReport(
                    "BudgetRollup",
                    new NamedRangeScope.Workbook(),
                    "SUM(Budget!$B$2:$B$3)",
                    GridGrindResponse.NamedRangeBackingKind.FORMULA)));
    GridGrindResponse.FormulaSurfaceReport formulaSurface =
        new GridGrindResponse.FormulaSurfaceReport(
            1,
            List.of(
                new GridGrindResponse.SheetFormulaSurfaceReport(
                    "Budget",
                    1,
                    1,
                    List.of(
                        new GridGrindResponse.FormulaPatternReport(
                            "SUM(B2:B3)", 1, List.of("B4"))))));
    GridGrindResponse.SheetSchemaReport schema =
        new GridGrindResponse.SheetSchemaReport(
            "Budget",
            "A1",
            3,
            2,
            2,
            List.of(
                new GridGrindResponse.SchemaColumnReport(
                    0,
                    "A1",
                    "Item",
                    2,
                    0,
                    List.of(new GridGrindResponse.TypeCountReport("STRING", 2)),
                    "STRING")));
    GridGrindResponse.FormulaHealthReport formulaHealth =
        new GridGrindResponse.FormulaHealthReport(
            1,
            new GridGrindResponse.AnalysisSummaryReport(1, 0, 0, 1),
            List.of(
                new GridGrindResponse.AnalysisFindingReport(
                    AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                    AnalysisSeverity.INFO,
                    "Volatile formula",
                    "Formula uses NOW().",
                    new GridGrindResponse.AnalysisLocationReport.Cell("Budget", "B4"),
                    List.of("NOW()"))));

    assertNull(formulaReport.target());
    assertEquals(1, surface.formulaBackedCount());
    assertEquals(1, formulaSurface.totalFormulaCellCount());
    assertEquals("STRING", schema.columns().getFirst().dominantType());
    assertEquals(1, formulaHealth.summary().infoCount());
    assertEquals(
        "Budget",
        ((GridGrindResponse.AnalysisLocationReport.Cell)
                formulaHealth.findings().getFirst().location())
            .sheetName());
  }

  @Test
  void problemContextDefaultNullsReturnNullForNonApplicableSubtypes() {
    GridGrindResponse.ProblemContext.ParseArguments parseArguments =
        new GridGrindResponse.ProblemContext.ParseArguments("--help");
    GridGrindResponse.ProblemContext.ValidateRequest validateRequest =
        new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE");

    assertNull(parseArguments.sourceType());
    assertNull(parseArguments.persistenceType());
    assertNull(parseArguments.requestPath());
    assertNull(parseArguments.jsonPath());
    assertNull(parseArguments.responsePath());
    assertNull(parseArguments.sourceWorkbookPath());
    assertNull(parseArguments.persistencePath());
    assertNull(parseArguments.operationIndex());
    assertNull(parseArguments.operationType());
    assertNull(parseArguments.readIndex());
    assertNull(parseArguments.readType());
    assertNull(parseArguments.requestId());
    assertNull(parseArguments.sheetName());
    assertNull(parseArguments.address());
    assertNull(parseArguments.range());
    assertNull(parseArguments.formula());
    assertNull(parseArguments.namedRangeName());
    assertNull(parseArguments.jsonLine());
    assertNull(parseArguments.jsonColumn());
    assertEquals("--help", parseArguments.argument());
    assertNull(validateRequest.argument());
  }

  @Test
  void responseRecordsRejectBlankAndInvalidValues() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WorkbookSummary.Empty(0, List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WorkbookSummary.Empty(1, List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WorkbookSummary.Empty(-1, List.of(), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WorkbookSummary.Empty(2, List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.Empty(0, List.of("Budget", "Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WorkbookSummary.Empty(1, List.of(" "), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Archive", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Archive"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), " ", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.WithSheets(
                0, List.of(), "Budget", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of(), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget", "Budget"), 0, false));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindResponse.SheetProtectionReport.Protected(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.PersistenceOutcome.SavedAs(" ", "/tmp/budget.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.PersistenceOutcome.SavedAs("budget.xlsx", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.SheetSummaryReport(
                " ",
                ExcelSheetVisibility.VISIBLE,
                new GridGrindResponse.SheetProtectionReport.Unprotected(),
                0,
                0,
                0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WindowReport("Budget", " ", 1, 1, List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.MergedRegionReport(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CellHyperlinkReport(
                " ", new HyperlinkTarget.Url("https://example.com")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CellCommentReport(
                " ", new GridGrindResponse.CommentReport("Review", "GridGrind", false)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.SheetLayoutReport(
                " ",
                new PaneReport.None(),
                100,
                SheetPresentationReport.defaults(),
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.ColumnLayoutReport(-1, 12.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.ColumnLayoutReport(0, 12.0, false, -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.ColumnLayoutReport(0, Double.NaN, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.ColumnLayoutReport(0, 0.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.ColumnLayoutReport(0, Double.POSITIVE_INFINITY, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.RowLayoutReport(-1, 12.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.RowLayoutReport(0, 12.0, false, -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.RowLayoutReport(0, 0.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.RowLayoutReport(0, Double.POSITIVE_INFINITY, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.FormulaSurfaceReport(-1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetFormulaSurfaceReport(" ", 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetFormulaSurfaceReport("Budget", -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetFormulaSurfaceReport("Budget", 0, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.FormulaPatternReport(" ", 1, List.of("A1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.FormulaPatternReport("SUM(A1)", 0, List.of("A1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetSchemaReport("Budget", " ", 1, 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WindowReport(" ", "A1", 1, 1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WindowReport("Budget", "A1", 0, 1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WindowReport("Budget", "A1", 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetSchemaReport(" ", "A1", 1, 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetSchemaReport("Budget", "A1", 0, 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetSchemaReport("Budget", "A1", 1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetSchemaReport("Budget", "A1", 1, 1, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SchemaColumnReport(-1, "A1", "Header", 0, 0, List.of(), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SchemaColumnReport(0, " ", "Header", 0, 0, List.of(), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SchemaColumnReport(0, "A1", "Header", -1, 0, List.of(), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SchemaColumnReport(0, "A1", "Header", 0, -1, List.of(), null));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.TypeCountReport(" ", 1));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.TypeCountReport("STRING", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.NamedRangeSurfaceReport(-1, 0, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.NamedRangeSurfaceReport(0, -1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.NamedRangeSurfaceReport(0, 0, -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.NamedRangeSurfaceReport(0, 0, 0, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.NamedRangeSurfaceEntryReport(
                " ",
                new NamedRangeScope.Workbook(),
                "Budget!$B$4",
                GridGrindResponse.NamedRangeBackingKind.RANGE));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.PersistenceOutcome.Overwritten(" ", "/tmp/budget.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.PersistenceOutcome.Overwritten("budget.xlsx", " "));
  }

  @Test
  void analysisReportsValidateAndCopyCollections() {
    List<String> evidence = new ArrayList<>(List.of("NOW()"));
    GridGrindResponse.AnalysisFindingReport workbookFinding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            AnalysisSeverity.INFO,
            "Volatile formula",
            "Formula uses NOW().",
            new GridGrindResponse.AnalysisLocationReport.Workbook(),
            evidence);
    GridGrindResponse.AnalysisFindingReport sheetFinding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_EXTERNAL_REFERENCE,
            AnalysisSeverity.WARNING,
            "External reference",
            "Formula references another workbook.",
            new GridGrindResponse.AnalysisLocationReport.Sheet("Budget"),
            List.of("[Book.xlsx]Sheet1!A1"));
    GridGrindResponse.AnalysisFindingReport rangeFinding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
            AnalysisSeverity.ERROR,
            "Broken named range",
            "Named range contains #REF!.",
            new GridGrindResponse.AnalysisLocationReport.Range("Budget", "A1:B2"),
            List.of("#REF!"));
    GridGrindResponse.AnalysisFindingReport namedRangeFinding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.NAMED_RANGE_SCOPE_SHADOWING,
            AnalysisSeverity.INFO,
            "Scope shadowing",
            "Name exists in multiple scopes.",
            new GridGrindResponse.AnalysisLocationReport.NamedRange(
                "BudgetTotal", new NamedRangeScope.Sheet("Budget")),
            List.of("Budget!$B$4"));
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(4, 1, 1, 2);
    GridGrindResponse.HyperlinkHealthReport hyperlinkHealth =
        new GridGrindResponse.HyperlinkHealthReport(
            2, summary, List.of(workbookFinding, sheetFinding));
    GridGrindResponse.NamedRangeHealthReport namedRangeHealth =
        new GridGrindResponse.NamedRangeHealthReport(2, summary, List.of(rangeFinding));
    GridGrindResponse.WorkbookFindingsReport workbookFindings =
        new GridGrindResponse.WorkbookFindingsReport(summary, List.of(namedRangeFinding));
    evidence.clear();

    assertEquals(1, hyperlinkHealth.findings().getFirst().evidence().size());
    assertEquals(
        "Budget",
        ((GridGrindResponse.AnalysisLocationReport.Sheet) sheetFinding.location()).sheetName());
    assertEquals(
        "A1:B2",
        ((GridGrindResponse.AnalysisLocationReport.Range) rangeFinding.location()).range());
    assertEquals(
        "BudgetTotal",
        ((GridGrindResponse.AnalysisLocationReport.NamedRange) namedRangeFinding.location())
            .name());
    assertEquals(1, namedRangeHealth.findings().size());
    assertEquals(1, workbookFindings.findings().size());

    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisSummaryReport(-1, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisSummaryReport(1, -1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisSummaryReport(1, 1, -1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisSummaryReport(1, 1, 1, -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisSummaryReport(1, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisLocationReport.Sheet(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisLocationReport.Cell("Budget", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisLocationReport.Cell(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisLocationReport.Range(" ", "A1:B2"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.AnalysisLocationReport.Range("Budget", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.AnalysisLocationReport.NamedRange(
                " ", new NamedRangeScope.Workbook()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.AnalysisFindingReport(
                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                AnalysisSeverity.INFO,
                " ",
                "message",
                new GridGrindResponse.AnalysisLocationReport.Workbook(),
                List.of("NOW()")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.AnalysisFindingReport(
                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                AnalysisSeverity.INFO,
                "title",
                " ",
                new GridGrindResponse.AnalysisLocationReport.Workbook(),
                List.of("NOW()")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.FormulaHealthReport(-1, summary, List.of(workbookFinding)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.HyperlinkHealthReport(-1, summary, List.of(workbookFinding)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.NamedRangeHealthReport(-1, summary, List.of(workbookFinding)));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindResponse.WorkbookFindingsReport(null, List.of(workbookFinding)));
  }

  @Test
  void paneReportVariantsAndProblemCauseCopyPathsRemainUsable() {
    PaneReport.None none = new PaneReport.None();
    PaneReport.Split split = new PaneReport.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT);
    assertInstanceOf(PaneReport.None.class, none);
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, split.activePane());

    GridGrindResponse.Problem problem =
        new GridGrindResponse.Problem(
            GridGrindProblemCode.INVALID_REQUEST,
            GridGrindProblemCategory.REQUEST,
            GridGrindProblemRecovery.CHANGE_REQUEST,
            "Invalid request",
            "bad request",
            "Fix the request.",
            new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"),
            null);
    assertEquals(List.of(), problem.causes());

    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.Problem(
                GridGrindProblemCode.INVALID_REQUEST,
                GridGrindProblemCategory.REQUEST,
                GridGrindProblemRecovery.CHANGE_REQUEST,
                "Invalid request",
                "bad request",
                "Fix the request.",
                new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"),
                java.util.Arrays.asList((GridGrindResponse.ProblemCause) null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.ProblemCause(GridGrindProblemCode.INVALID_REQUEST, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.ProblemCause(
                GridGrindProblemCode.INVALID_REQUEST, "bad request", " "));
  }

  @Test
  void validatesCellFillReportContracts() {
    CellFillReport noFill =
        new CellFillReport(dev.erst.gridgrind.excel.ExcelFillPattern.NONE, null, null);
    CellFillReport patternedFill =
        new CellFillReport(dev.erst.gridgrind.excel.ExcelFillPattern.BRICKS, null, rgb("#aabbcc"));

    assertEquals(dev.erst.gridgrind.excel.ExcelFillPattern.NONE, noFill.pattern());
    assertNull(noFill.foregroundColor());
    assertNull(noFill.backgroundColor());
    assertEquals(rgb("#AABBCC"), patternedFill.backgroundColor());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillReport(
                dev.erst.gridgrind.excel.ExcelFillPattern.NONE, rgb("#112233"), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillReport(
                dev.erst.gridgrind.excel.ExcelFillPattern.NONE, null, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillReport(
                dev.erst.gridgrind.excel.ExcelFillPattern.NONE, rgb("#112233"), rgb("#445566")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillReport(
                dev.erst.gridgrind.excel.ExcelFillPattern.SOLID, rgb("#112233"), rgb("#445566")));
  }

  private static GridGrindResponse.CellStyleReport defaultCellStyleReport() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        new CellAlignmentReport(
            true, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP, 0, 0),
        new CellFontReport(
            true,
            false,
            "Aptos",
            new FontHeightReport(230, new BigDecimal("11.5")),
            rgb("#1F4E78"),
            true,
            false),
        new CellFillReport(dev.erst.gridgrind.excel.ExcelFillPattern.SOLID, rgb("#FFF2CC"), null),
        new CellBorderReport(
            new CellBorderSideReport(ExcelBorderStyle.THIN, null),
            new CellBorderSideReport(ExcelBorderStyle.DOUBLE, null),
            new CellBorderSideReport(ExcelBorderStyle.THIN, null),
            new CellBorderSideReport(ExcelBorderStyle.THIN, null)),
        new CellProtectionReport(true, false));
  }

  private static CellColorReport rgb(String rgb) {
    return new CellColorReport(rgb);
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }
}
