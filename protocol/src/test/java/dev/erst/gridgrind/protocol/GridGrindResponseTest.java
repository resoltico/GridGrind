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
    List<GridGrindResponse.SheetReport> sheets = new ArrayList<>();
    sheets.add(new GridGrindResponse.SheetReport("Budget", 1, 0, 1, List.of(), List.of()));
    List<GridGrindResponse.NamedRangeReport> namedRanges =
        new ArrayList<>(
            List.of(
                new GridGrindResponse.NamedRangeReport.RangeReport(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    "Budget!$B$4",
                    new NamedRangeTarget("Budget", "B4"))));

    GridGrindResponse.Success success =
        new GridGrindResponse.Success(
            null,
            "/tmp/budget.xlsx",
            new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, true),
            namedRanges,
            sheets);
    sheets.clear();
    namedRanges.clear();

    assertInstanceOf(GridGrindResponse.Success.class, success);
    assertEquals(GridGrindProtocolVersion.V1, success.protocolVersion());
    assertEquals("/tmp/budget.xlsx", success.savedWorkbookPath());
    assertEquals(List.of("Budget"), success.workbook().sheetNames());
    assertEquals(1, success.workbook().namedRangeCount());
    assertEquals(1, success.namedRanges().size());
    assertEquals(1, success.sheets().size());

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
    List<GridGrindResponse.NamedRangeReport> namedRanges =
        new ArrayList<>(
            List.of(
                new GridGrindResponse.NamedRangeReport.FormulaReport(
                    "BudgetTotal", new NamedRangeScope.Workbook(), "SUM(Budget!$B$2:$B$3)")));
    GridGrindResponse.CellStyleReport style = defaultCellStyleReport();
    List<GridGrindResponse.CellReport> cells =
        new ArrayList<>(List.of(textReport("A1", "Item", style)));
    List<GridGrindResponse.PreviewRowReport> previewRows =
        new ArrayList<>(List.of(new GridGrindResponse.PreviewRowReport(0, cells)));

    GridGrindResponse.WorkbookSummary workbook =
        new GridGrindResponse.WorkbookSummary(1, sheetNames, 1, true);
    GridGrindResponse.SheetReport report =
        new GridGrindResponse.SheetReport("Budget", 1, 0, 0, cells, previewRows);
    GridGrindResponse.Success success =
        new GridGrindResponse.Success(null, null, workbook, namedRanges, List.of(report));

    sheetNames.clear();
    namedRanges.clear();
    cells.clear();
    previewRows.clear();

    assertEquals(List.of("Budget"), workbook.sheetNames());
    assertEquals(1, success.namedRanges().size());
    assertEquals(1, report.requestedCells().size());
    assertEquals(1, report.previewRows().size());

    assertThrows(
        NullPointerException.class,
        () -> new GridGrindResponse.Success(null, null, null, null, null));
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

    GridGrindResponse.Success emptySuccess =
        new GridGrindResponse.Success(
            null, null, new GridGrindResponse.WorkbookSummary(0, List.of(), 0, false), null, null);
    assertEquals(List.of(), emptySuccess.namedRanges());
    assertEquals(List.of(), emptySuccess.sheets());
  }

  @Test
  void problemContextRecordsReturnCorrectStages() {
    GridGrindResponse.ProblemContext.ParseArguments parseArgs =
        new GridGrindResponse.ProblemContext.ParseArguments("--response");
    assertEquals("PARSE_ARGUMENTS", parseArgs.stage());
    assertEquals("--response", parseArgs.argument());

    GridGrindResponse.ProblemContext.ReadRequest readRequest =
        new GridGrindResponse.ProblemContext.ReadRequest(
            "/tmp/request.json", "analysis.sheets[0]", 7, 21);
    assertEquals("READ_REQUEST", readRequest.stage());
    assertEquals("/tmp/request.json", readRequest.requestPath());
    assertNull(new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE").argument());

    GridGrindResponse.ProblemContext.ApplyOperation applyOperation =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            "NEW", "SAVE_AS", 1, "SET_NAMED_RANGE", "Budget", "B4", "B4", null, "BudgetTotal");
    assertEquals("APPLY_OPERATION", applyOperation.stage());
    assertEquals("BudgetTotal", applyOperation.namedRangeName());

    GridGrindResponse.ProblemContext.AnalyzeWorkbook analyzeWorkbook =
        new GridGrindResponse.ProblemContext.AnalyzeWorkbook(
            "NEW", "SAVE_AS", "Budget", "B4", "SUM(B2:B3)", "BudgetTotal");
    assertEquals("ANALYZE_WORKBOOK", analyzeWorkbook.stage());
    assertEquals("BudgetTotal", analyzeWorkbook.namedRangeName());
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

    GridGrindResponse.ProblemContext.AnalyzeWorkbook analyzeBase =
        new GridGrindResponse.ProblemContext.AnalyzeWorkbook("NEW", "NONE", null, null, null, null);
    GridGrindResponse.ProblemContext.AnalyzeWorkbook enrichedAnalyze =
        analyzeBase.withExceptionData("Budget", "C1", "SUM(B1:B3)", "BudgetTotal");

    assertEquals("Budget", enrichedAnalyze.sheetName());
    assertEquals("C1", enrichedAnalyze.address());
    assertEquals("SUM(B1:B3)", enrichedAnalyze.formula());
    assertEquals("BudgetTotal", enrichedAnalyze.namedRangeName());

    GridGrindResponse.ProblemContext.AnalyzeWorkbook populatedAnalyze =
        new GridGrindResponse.ProblemContext.AnalyzeWorkbook(
            "NEW", "NONE", "Sheet1", "A1", "SUM(B1:B2)", "ExistingRange");
    GridGrindResponse.ProblemContext.AnalyzeWorkbook preservedAnalyze =
        populatedAnalyze.withExceptionData("OtherSheet", "B2", "OTHER()", "OtherRange");
    assertEquals("Sheet1", preservedAnalyze.sheetName());
    assertEquals("A1", preservedAnalyze.address());
    assertEquals("SUM(B1:B2)", preservedAnalyze.formula());
    assertEquals("ExistingRange", preservedAnalyze.namedRangeName());
  }

  @Test
  void readRequestWithJsonPopulatesAndPreservesJsonLocationFields() {
    GridGrindResponse.ProblemContext.ReadRequest blank =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);
    GridGrindResponse.ProblemContext.ReadRequest enriched = blank.withJson("$.analysis", 4, 12);

    assertEquals("$.analysis", enriched.jsonPath());
    assertEquals(4, enriched.jsonLine());
    assertEquals(12, enriched.jsonColumn());

    GridGrindResponse.ProblemContext.ReadRequest populated =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", "$.existing", 1, 2);
    GridGrindResponse.ProblemContext.ReadRequest preserved = populated.withJson("$.other", 9, 10);
    assertEquals("$.existing", preserved.jsonPath());
    assertEquals(1, preserved.jsonLine());
    assertEquals(2, preserved.jsonColumn());
  }

  @Test
  void cellReportSubtypesExposeTypedGettersAndMetadataDefaults() {
    GridGrindResponse.CellStyleReport style = defaultCellStyleReport();
    GridGrindResponse.HyperlinkReport hyperlink =
        new GridGrindResponse.HyperlinkReport(ExcelHyperlinkType.URL, "https://example.com");
    GridGrindResponse.CommentReport comment =
        new GridGrindResponse.CommentReport("Review", "GridGrind", false);

    GridGrindResponse.CellReport.BlankReport blankReport =
        new GridGrindResponse.CellReport.BlankReport("A1", "BLANK", "", style, null, null);
    GridGrindResponse.CellReport.TextReport textReport =
        new GridGrindResponse.CellReport.TextReport(
            "A2", "STRING", "Hello", style, hyperlink, comment, "Hello");
    GridGrindResponse.CellReport.NumberReport numberReport =
        new GridGrindResponse.CellReport.NumberReport(
            "B1", "NUMERIC", "42", style, null, null, 42.0);
    GridGrindResponse.CellReport.BooleanReport booleanReport =
        new GridGrindResponse.CellReport.BooleanReport(
            "C1", "BOOLEAN", "TRUE", style, null, null, true);
    GridGrindResponse.CellReport.ErrorReport errorReport =
        new GridGrindResponse.CellReport.ErrorReport(
            "D1", "ERROR", "#DIV/0!", style, null, null, "#DIV/0!");
    GridGrindResponse.CellReport.FormulaReport formulaReport =
        new GridGrindResponse.CellReport.FormulaReport(
            "E1",
            "FORMULA",
            "85",
            style,
            null,
            null,
            "SUM(B1:B2)",
            new GridGrindResponse.CellReport.NumberReport(
                "E1", "NUMERIC", "85", style, null, null, 85.0));

    assertEquals("BLANK", blankReport.effectiveType());
    assertEquals("STRING", textReport.effectiveType());
    assertEquals("NUMERIC", numberReport.effectiveType());
    assertEquals("BOOLEAN", booleanReport.effectiveType());
    assertEquals("ERROR", errorReport.effectiveType());
    assertEquals("FORMULA", formulaReport.effectiveType());
    assertEquals(ExcelHyperlinkType.URL, textReport.hyperlink().type());
    assertEquals("Review", textReport.comment().text());
    assertNull(blankReport.formula());
    assertNull(blankReport.hyperlink());
    assertNull(blankReport.comment());
    assertNull(blankReport.stringValue());
    assertNull(blankReport.numberValue());
    assertNull(blankReport.booleanValue());
    assertNull(blankReport.errorValue());
    assertNull(numberReport.stringValue());
    assertNull(booleanReport.numberValue());
    assertNull(errorReport.booleanValue());
    assertNull(formulaReport.errorValue());
  }

  @Test
  void namedRangeFormulaReportReturnsNullTargetByDefault() {
    GridGrindResponse.NamedRangeReport.FormulaReport formulaReport =
        new GridGrindResponse.NamedRangeReport.FormulaReport(
            "BudgetRollup", new NamedRangeScope.Workbook(), "SUM(Budget!$B$2:$B$3)");

    assertNull(formulaReport.target());
  }

  @Test
  void problemContextDefaultNullsReturnNullForNonApplicableSubtypes() {
    GridGrindResponse.ProblemContext.ParseArguments parseArgs =
        new GridGrindResponse.ProblemContext.ParseArguments("--help");

    assertNull(parseArgs.sourceMode());
    assertNull(parseArgs.persistenceMode());
    assertNull(parseArgs.requestPath());
    assertNull(parseArgs.jsonPath());
    assertNull(parseArgs.jsonLine());
    assertNull(parseArgs.jsonColumn());
    assertNull(parseArgs.responsePath());
    assertNull(parseArgs.sourceWorkbookPath());
    assertNull(parseArgs.persistencePath());
    assertNull(parseArgs.operationIndex());
    assertNull(parseArgs.operationType());
    assertNull(parseArgs.sheetName());
    assertNull(parseArgs.address());
    assertNull(parseArgs.range());
    assertNull(parseArgs.formula());
    assertNull(parseArgs.namedRangeName());
    assertEquals("--help", parseArgs.argument());
  }

  @Test
  void problemCopiesNullCauseListsToEmptyLists() {
    GridGrindResponse.Problem problem =
        new GridGrindResponse.Problem(
            GridGrindProblemCode.INVALID_REQUEST,
            GridGrindProblemCategory.REQUEST,
            GridGrindProblemRecovery.CHANGE_REQUEST,
            "Invalid request",
            "bad request",
            "Fix the request and retry.",
            new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"),
            null);

    assertEquals(List.of(), problem.causes());
  }

  private GridGrindResponse.CellReport.TextReport textReport(
      String address, String value, GridGrindResponse.CellStyleReport style) {
    return new GridGrindResponse.CellReport.TextReport(
        address, "STRING", value, style, null, null, value);
  }

  private GridGrindResponse.CellStyleReport defaultCellStyleReport() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        false,
        false,
        false,
        ExcelHorizontalAlignment.GENERAL,
        ExcelVerticalAlignment.BOTTOM,
        "Calibri",
        new FontHeightReport(220, new BigDecimal("11")),
        null,
        false,
        false,
        null,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE);
  }
}
