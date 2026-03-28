package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlinkType;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
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
            "workbook", new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, true)));

    GridGrindResponse.Success success =
        new GridGrindResponse.Success(
            null,
            new GridGrindResponse.PersistenceOutcome.SavedAs("budget.xlsx", "/tmp/budget.xlsx"),
            reads);
    reads.clear();

    assertInstanceOf(GridGrindResponse.Success.class, success);
    assertEquals(GridGrindProtocolVersion.V1, success.protocolVersion());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.SavedAs.class, success.persistence());
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
        new GridGrindResponse.WorkbookSummary(1, sheetNames, 1, true);
    List<WorkbookReadResult> reads =
        new ArrayList<>(
            List.of(new WorkbookReadResult.WorkbookSummaryResult("workbook", workbook)));

    GridGrindResponse.Success success = new GridGrindResponse.Success(null, null, reads);
    sheetNames.clear();
    reads.clear();

    assertEquals(List.of("Budget"), workbook.sheetNames());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, success.persistence());
    assertEquals(1, success.reads().size());

    assertThrows(NullPointerException.class, () -> new GridGrindResponse.Success(null, null, null));
    assertThrows(
        NullPointerException.class, () -> new GridGrindResponse.WorkbookSummary(1, null, 0, true));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.CellReport.TextReport(
                "A1", "STRING", "x", null, null, null, "x"));
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
    GridGrindResponse.HyperlinkReport hyperlink =
        new GridGrindResponse.HyperlinkReport(ExcelHyperlinkType.URL, "https://example.com");
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
                "E1", "NUMERIC", "85", style, null, null, 85.0));

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
            new GridGrindResponse.FreezePaneReport.Frozen(1, 1, 1, 1),
            List.of(new GridGrindResponse.ColumnLayoutReport(0, 12.5d)),
            List.of(new GridGrindResponse.RowLayoutReport(0, 18.0d)));

    assertEquals("FORMULA", formulaReport.effectiveType());
    assertEquals(
        "STRING",
        new GridGrindResponse.CellReport.TextReport("B1", "STRING", "x", style, null, null, "x")
            .effectiveType());
    assertEquals(
        "NUMERIC",
        new GridGrindResponse.CellReport.NumberReport(
                "B2", "NUMERIC", "12", style, null, null, 12.0)
            .effectiveType());
    assertEquals(
        "BOOLEAN",
        new GridGrindResponse.CellReport.BooleanReport(
                "B3", "BOOLEAN", "TRUE", style, null, null, true)
            .effectiveType());
    assertEquals(ExcelHyperlinkType.URL, formulaReport.hyperlink().type());
    assertEquals("Review", formulaReport.comment().text());
    assertNull(blankReport.formula());
    assertNull(blankReport.hyperlink());
    assertNull(blankReport.comment());
    assertNull(blankReport.stringValue());
    assertNull(blankReport.numberValue());
    assertNull(blankReport.booleanValue());
    assertNull(blankReport.errorValue());
    assertEquals("A1", window.rows().getFirst().cells().getFirst().address());
    assertInstanceOf(GridGrindResponse.FreezePaneReport.Frozen.class, layout.freezePanes());
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
                    dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                        .FORMULA_VOLATILE_FUNCTION,
                    dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
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
        () -> new GridGrindResponse.PersistenceOutcome.SavedAs(" ", "/tmp/budget.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.PersistenceOutcome.SavedAs("budget.xlsx", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.SheetSummaryReport(" ", 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.WindowReport("Budget", " ", 1, 1, List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.MergedRegionReport(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CellHyperlinkReport(
                " ",
                new GridGrindResponse.HyperlinkReport(
                    ExcelHyperlinkType.URL, "https://example.com")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CellCommentReport(
                " ", new GridGrindResponse.CommentReport("Review", "GridGrind", false)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.SheetLayoutReport(
                " ", new GridGrindResponse.FreezePaneReport.None(), List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.ColumnLayoutReport(-1, 12.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.ColumnLayoutReport(0, Double.NaN));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.ColumnLayoutReport(0, 0.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.ColumnLayoutReport(0, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.RowLayoutReport(-1, 12.0));
    assertThrows(
        IllegalArgumentException.class, () -> new GridGrindResponse.RowLayoutReport(0, 0.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindResponse.RowLayoutReport(0, Double.POSITIVE_INFINITY));
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
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
            "Volatile formula",
            "Formula uses NOW().",
            new GridGrindResponse.AnalysisLocationReport.Workbook(),
            evidence);
    GridGrindResponse.AnalysisFindingReport sheetFinding =
        new GridGrindResponse.AnalysisFindingReport(
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                .FORMULA_EXTERNAL_REFERENCE,
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.WARNING,
            "External reference",
            "Formula references another workbook.",
            new GridGrindResponse.AnalysisLocationReport.Sheet("Budget"),
            List.of("[Book.xlsx]Sheet1!A1"));
    GridGrindResponse.AnalysisFindingReport rangeFinding =
        new GridGrindResponse.AnalysisFindingReport(
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                .NAMED_RANGE_BROKEN_REFERENCE,
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.ERROR,
            "Broken named range",
            "Named range contains #REF!.",
            new GridGrindResponse.AnalysisLocationReport.Range("Budget", "A1:B2"),
            List.of("#REF!"));
    GridGrindResponse.AnalysisFindingReport namedRangeFinding =
        new GridGrindResponse.AnalysisFindingReport(
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                .NAMED_RANGE_SCOPE_SHADOWING,
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
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
                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                    .FORMULA_VOLATILE_FUNCTION,
                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
                " ",
                "message",
                new GridGrindResponse.AnalysisLocationReport.Workbook(),
                List.of("NOW()")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.AnalysisFindingReport(
                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                    .FORMULA_VOLATILE_FUNCTION,
                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
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
  void freezePaneNoneAndProblemCauseCopyPathsRemainUsable() {
    GridGrindResponse.FreezePaneReport.None none = new GridGrindResponse.FreezePaneReport.None();
    assertInstanceOf(GridGrindResponse.FreezePaneReport.None.class, none);

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
  }

  private static GridGrindResponse.CellStyleReport defaultCellStyleReport() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        true,
        false,
        true,
        ExcelHorizontalAlignment.CENTER,
        ExcelVerticalAlignment.TOP,
        "Aptos",
        new FontHeightReport(230, new BigDecimal("11.5")),
        "#1F4E78",
        true,
        false,
        "#FFF2CC",
        ExcelBorderStyle.THIN,
        ExcelBorderStyle.DOUBLE,
        ExcelBorderStyle.THIN,
        ExcelBorderStyle.THIN);
  }
}
