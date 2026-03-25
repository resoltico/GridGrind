package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for GridGrindResponse sealed interface record construction and validation. */
class GridGrindResponseTest {
  @Test
  void createsSuccessAndFailureResponses() {
    List<GridGrindResponse.SheetReport> reports = new ArrayList<>();
    reports.add(new GridGrindResponse.SheetReport("Budget", 1, 0, 1, List.of(), List.of()));

    GridGrindResponse.Success success =
        new GridGrindResponse.Success(
            null,
            "/tmp/budget.xlsx",
            new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), true),
            reports);
    reports.clear();

    assertInstanceOf(GridGrindResponse.Success.class, success);
    assertEquals(GridGrindProtocolVersion.V1, success.protocolVersion());
    assertEquals("/tmp/budget.xlsx", success.savedWorkbookPath());
    assertEquals(List.of("Budget"), success.workbook().sheetNames());
    assertEquals(1, success.sheets().size());

    GridGrindResponse.Failure failure =
        new GridGrindResponse.Failure(
            null,
            GridGrindResponse.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "Bad request",
                new GridGrindResponse.ProblemContext.ValidateRequest(null, null)));

    assertInstanceOf(GridGrindResponse.Failure.class, failure);
    assertEquals(GridGrindProtocolVersion.V1, failure.protocolVersion());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals(GridGrindProblemCategory.REQUEST, failure.problem().category());
    assertEquals(GridGrindProblemRecovery.CHANGE_REQUEST, failure.problem().recovery());
    assertEquals("Invalid request", failure.problem().title());
    assertEquals("Bad request", failure.problem().message());
    assertEquals(List.of(), failure.problem().causes());
  }

  @Test
  void copiesAndValidatesNestedCollections() {
    List<String> sheetNames = new ArrayList<>(List.of("Budget"));
    GridGrindResponse.CellStyleReport cellStyle =
        new GridGrindResponse.CellStyleReport("General", false, false, false, "GENERAL", "BOTTOM");
    List<GridGrindResponse.CellReport> cells =
        new ArrayList<>(
            List.of(
                new GridGrindResponse.CellReport.TextReport(
                    "A1", "STRING", "Item", cellStyle, "Item")));
    List<GridGrindResponse.PreviewRowReport> previewRows =
        new ArrayList<>(List.of(new GridGrindResponse.PreviewRowReport(0, cells)));

    GridGrindResponse.WorkbookSummary workbook =
        new GridGrindResponse.WorkbookSummary(1, sheetNames, true);
    GridGrindResponse.SheetReport report =
        new GridGrindResponse.SheetReport("Budget", 1, 0, 0, cells, previewRows);

    sheetNames.clear();
    cells.clear();
    previewRows.clear();

    assertEquals(List.of("Budget"), workbook.sheetNames());
    assertEquals(1, report.requestedCells().size());
    assertEquals(1, report.previewRows().size());
    assertEquals("General", report.requestedCells().getFirst().style().numberFormat());

    assertThrows(
        NullPointerException.class, () -> new GridGrindResponse.Success(null, null, null, null));
    assertThrows(
        NullPointerException.class, () -> new GridGrindResponse.WorkbookSummary(1, null, true));
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindResponse.CellReport.TextReport("A1", "STRING", "x", null, "x"));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.CellStyleReport(null, false, false, false, "GENERAL", "BOTTOM"));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.Problem(
                null,
                GridGrindProblemCategory.REQUEST,
                GridGrindProblemRecovery.CHANGE_REQUEST,
                "title",
                "message",
                "resolution",
                new GridGrindResponse.ProblemContext.ExecuteRequest(null, null),
                List.of()));

    GridGrindResponse.Success emptySuccess =
        new GridGrindResponse.Success(
            null, null, new GridGrindResponse.WorkbookSummary(0, List.of(), false), null);
    assertEquals(GridGrindProtocolVersion.V1, emptySuccess.protocolVersion());
    assertEquals(List.of(), emptySuccess.sheets());

    GridGrindResponse.Success explicitNullVersion =
        new GridGrindResponse.Success(
            null, null, new GridGrindResponse.WorkbookSummary(0, List.of(), false), List.of());
    assertEquals(GridGrindProtocolVersion.V1, explicitNullVersion.protocolVersion());
  }

  @Test
  void problemContextRecordsReturnCorrectStages() {
    GridGrindResponse.ProblemContext parseArgs =
        new GridGrindResponse.ProblemContext.ParseArguments("--response");
    assertEquals("PARSE_ARGUMENTS", parseArgs.stage());
    assertEquals("--response", parseArgs.argument());

    GridGrindResponse.ProblemContext readRequest =
        new GridGrindResponse.ProblemContext.ReadRequest(
            "/tmp/request.json", "analysis.sheets[0]", 7, 21);
    assertEquals("READ_REQUEST", readRequest.stage());
    assertEquals("/tmp/request.json", readRequest.requestPath());
    assertEquals("analysis.sheets[0]", readRequest.jsonPath());
    assertEquals(7, readRequest.jsonLine());
    assertEquals(21, readRequest.jsonColumn());

    GridGrindResponse.ProblemContext validateRequest =
        new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "SAVE_AS");
    assertEquals("VALIDATE_REQUEST", validateRequest.stage());
    assertEquals("NEW", validateRequest.sourceMode());
    assertEquals("SAVE_AS", validateRequest.persistenceMode());

    GridGrindResponse.ProblemContext openWorkbook =
        new GridGrindResponse.ProblemContext.OpenWorkbook("NEW", "SAVE_AS", "/tmp/source.xlsx");
    assertEquals("OPEN_WORKBOOK", openWorkbook.stage());
    assertEquals("/tmp/source.xlsx", openWorkbook.sourceWorkbookPath());

    GridGrindResponse.ProblemContext applyOperation =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            "NEW", "SAVE_AS", 1, "SET_RANGE", "Budget", "B4", "A1:B4", "SUM(B2:B3)");
    assertEquals("APPLY_OPERATION", applyOperation.stage());
    assertEquals(1, applyOperation.operationIndex());
    assertEquals("SET_RANGE", applyOperation.operationType());
    assertEquals("Budget", applyOperation.sheetName());
    assertEquals("B4", applyOperation.address());
    assertEquals("A1:B4", applyOperation.range());
    assertEquals("SUM(B2:B3)", applyOperation.formula());

    GridGrindResponse.ProblemContext persistWorkbook =
        new GridGrindResponse.ProblemContext.PersistWorkbook(
            "NEW", "SAVE_AS", "/tmp/source.xlsx", "/tmp/output.xlsx");
    assertEquals("PERSIST_WORKBOOK", persistWorkbook.stage());
    assertEquals("/tmp/output.xlsx", persistWorkbook.persistencePath());

    GridGrindResponse.ProblemContext analyzeWorkbook =
        new GridGrindResponse.ProblemContext.AnalyzeWorkbook(
            "NEW", "SAVE_AS", "Budget", "B4", "SUM(B2:B3)");
    assertEquals("ANALYZE_WORKBOOK", analyzeWorkbook.stage());

    GridGrindResponse.ProblemContext executeRequest =
        new GridGrindResponse.ProblemContext.ExecuteRequest("NEW", "NONE");
    assertEquals("EXECUTE_REQUEST", executeRequest.stage());

    GridGrindResponse.ProblemContext writeResponse =
        new GridGrindResponse.ProblemContext.WriteResponse("/tmp/response.json");
    assertEquals("WRITE_RESPONSE", writeResponse.stage());
    assertEquals("/tmp/response.json", writeResponse.responsePath());
  }

  @Test
  void problemRecordIsFullyPopulated() {
    GridGrindResponse.Problem problem =
        new GridGrindResponse.Problem(
            GridGrindProblemCode.IO_ERROR,
            GridGrindProblemCategory.IO,
            GridGrindProblemRecovery.CHECK_ENVIRONMENT,
            "I/O failure",
            "disk failed",
            "Check file paths, permissions, file locks, and disk state before retrying.",
            new GridGrindResponse.ProblemContext.WriteResponse("/tmp/response.json"),
            null);

    GridGrindResponse.Failure response = new GridGrindResponse.Failure(null, problem);

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    assertEquals("/tmp/response.json", response.problem().context().responsePath());
    assertEquals(List.of(), response.problem().causes());
  }

  @Test
  void withJsonEnrichesReadRequestContext() {
    GridGrindResponse.ProblemContext.ReadRequest base =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);
    GridGrindResponse.ProblemContext.ReadRequest enriched =
        base.withJson("analysis.sheets[0]", 7, 21);

    assertEquals("/tmp/request.json", enriched.requestPath());
    assertEquals("analysis.sheets[0]", enriched.jsonPath());
    assertEquals(7, enriched.jsonLine());
    assertEquals(21, enriched.jsonColumn());
  }

  @Test
  void withExceptionDataEnrichesApplyOperationContext() {
    GridGrindResponse.ProblemContext.ApplyOperation base =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            "NEW", "NONE", 0, "SET_CELL", null, null, null, null);
    GridGrindResponse.ProblemContext.ApplyOperation enriched =
        base.withExceptionData("Budget", "B4", null, "SUM(B2:B3)");

    assertEquals("Budget", enriched.sheetName());
    assertEquals("B4", enriched.address());
    assertNull(enriched.range());
    assertEquals("SUM(B2:B3)", enriched.formula());
  }

  @Test
  void withExceptionDataEnrichesAnalyzeWorkbookContext() {
    GridGrindResponse.ProblemContext.AnalyzeWorkbook base =
        new GridGrindResponse.ProblemContext.AnalyzeWorkbook("NEW", "NONE", null, null, null);
    GridGrindResponse.ProblemContext.AnalyzeWorkbook enriched =
        base.withExceptionData("Budget", "C1", "SUM(B1:B3)");

    assertEquals("Budget", enriched.sheetName());
    assertEquals("C1", enriched.address());
    assertEquals("SUM(B1:B3)", enriched.formula());

    // Existing non-null fields are preserved (true branches of the ternary guards)
    GridGrindResponse.ProblemContext.AnalyzeWorkbook populated =
        new GridGrindResponse.ProblemContext.AnalyzeWorkbook(
            "NEW", "NONE", "Sheet1", "A1", "SUM(B1:B2)");
    GridGrindResponse.ProblemContext.AnalyzeWorkbook preservedPopulated =
        populated.withExceptionData("OtherSheet", "B2", "OTHER()");
    assertEquals("Sheet1", preservedPopulated.sheetName());
    assertEquals("A1", preservedPopulated.address());
    assertEquals("SUM(B1:B2)", preservedPopulated.formula());
  }

  @Test
  void cellReportSubtypesExposeTypedGettersAndDefaultNulls() {
    GridGrindResponse.CellStyleReport style =
        new GridGrindResponse.CellStyleReport("General", false, false, false, "GENERAL", "BOTTOM");

    GridGrindResponse.CellReport.BlankReport blankReport =
        new GridGrindResponse.CellReport.BlankReport("A1", "BLANK", "", style);
    GridGrindResponse.CellReport.TextReport textReport =
        new GridGrindResponse.CellReport.TextReport("A2", "STRING", "Hello", style, "Hello");
    GridGrindResponse.CellReport.NumberReport numberReport =
        new GridGrindResponse.CellReport.NumberReport("B1", "NUMERIC", "42", style, 42.0);
    GridGrindResponse.CellReport.BooleanReport booleanReport =
        new GridGrindResponse.CellReport.BooleanReport("C1", "BOOLEAN", "TRUE", style, true);
    GridGrindResponse.CellReport.ErrorReport errorReport =
        new GridGrindResponse.CellReport.ErrorReport("D1", "ERROR", "#DIV/0!", style, "#DIV/0!");
    GridGrindResponse.CellReport.FormulaReport formulaReport =
        new GridGrindResponse.CellReport.FormulaReport(
            "E1",
            "FORMULA",
            "85",
            style,
            "SUM(B1:B2)",
            new GridGrindResponse.CellReport.NumberReport("E1", "NUMERIC", "85", style, 85.0));

    assertEquals("BLANK", blankReport.effectiveType());
    assertEquals("STRING", textReport.effectiveType());
    assertEquals("NUMERIC", numberReport.effectiveType());
    assertEquals("BOOLEAN", booleanReport.effectiveType());
    assertEquals("ERROR", errorReport.effectiveType());
    assertEquals("FORMULA", formulaReport.effectiveType());

    // Default null methods return null for subtypes that do not populate them
    assertNull(blankReport.formula());
    assertNull(blankReport.stringValue());
    assertNull(blankReport.numberValue());
    assertNull(blankReport.booleanValue());
    assertNull(blankReport.errorValue());
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
    assertEquals("--help", parseArgs.argument());

    // argument() default returns null for subtypes that are not ParseArguments
    GridGrindResponse.ProblemContext.ValidateRequest validateRequest =
        new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE");
    assertNull(validateRequest.argument());
  }
}
