package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.time.DateTimeException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for GridGrindProblems exception classification and problem construction. */
class GridGrindProblemsTest {
  @Test
  void classifiesProblemCodesAcrossAllKnownFamilies() {
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND,
        GridGrindProblems.codeFor(new WorkbookNotFoundException(Path.of("missing.xlsx"))));
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND,
        GridGrindProblems.codeFor(new SheetNotFoundException("Budget")));
    assertEquals(
        GridGrindProblemCode.CELL_NOT_FOUND,
        GridGrindProblems.codeFor(new CellNotFoundException("A1")));
    assertEquals(
        GridGrindProblemCode.INVALID_CELL_ADDRESS,
        GridGrindProblems.codeFor(
            new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_RANGE_ADDRESS,
        GridGrindProblems.codeFor(
            new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.UNSUPPORTED_FORMULA,
        GridGrindProblems.codeFor(
            new UnsupportedFormulaException(
                "Budget",
                "C1",
                "LAMBDA(x,x+1)(2)",
                "unsupported",
                new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        GridGrindProblems.codeFor(
            new InvalidFormulaException(
                "Budget", "B4", "SUM(", "bad formula", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_JSON,
        GridGrindProblems.codeFor(
            new InvalidJsonException(
                "bad json", "analysis.sheets[0]", 4, 12, new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(
            new InvalidRequestException(
                "bad request",
                "analysis.sheets[0].previewRowCount",
                6,
                41,
                new IllegalArgumentException("bad"))));
    assertEquals(GridGrindProblemCode.IO_ERROR, GridGrindProblems.codeFor(new IOException("disk")));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(new IllegalArgumentException("bad")));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(new DateTimeException("bad date")));
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        GridGrindProblems.codeFor(new UnsupportedOperationException("boom")));
  }

  @Test
  void buildsStructuredProblemsAndCauseChains() {
    GridGrindResponse.ProblemContext context =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            null, null, null, null, "Budget", "B4", null, "SUM(");

    InvalidFormulaException exception =
        new InvalidFormulaException(
            "Budget", "B4", "SUM(", "bad formula", new IllegalArgumentException("root cause"));
    GridGrindResponse.Problem problem = GridGrindProblems.fromException(exception, context);

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, problem.code());
    assertEquals(GridGrindProblemCategory.FORMULA, problem.category());
    assertEquals(GridGrindProblemRecovery.CHANGE_REQUEST, problem.recovery());
    assertEquals("Invalid formula", problem.title());
    assertEquals(
        "Fix the formula syntax or workbook references, then retry.", problem.resolution());
    assertEquals(context, problem.context());
    assertEquals(2, problem.causes().size());
    assertEquals("InvalidFormulaException", problem.causes().get(0).type());
    assertEquals("root cause", problem.causes().get(1).message());
    assertEquals("SUM(", GridGrindProblems.formulaFor(exception));
    assertEquals("Budget", GridGrindProblems.sheetNameFor(exception));
    assertEquals("B4", GridGrindProblems.addressFor(exception));
    assertNull(GridGrindProblems.formulaFor(new IllegalArgumentException("bad")));
    assertNull(GridGrindProblems.sheetNameFor(new IllegalArgumentException("bad")));
    assertNull(GridGrindProblems.addressFor(new IllegalArgumentException("bad")));
    assertNull(GridGrindProblems.rangeFor(new IllegalArgumentException("bad")));
  }

  @Test
  void supportsExplicitProblemsAndSupplementalDiagnostics() {
    GridGrindResponse.ProblemContext context =
        new GridGrindResponse.ProblemContext.WriteResponse("/tmp/response.json");

    GridGrindResponse.Problem explicitWithoutCauses =
        GridGrindProblems.problem(
            GridGrindProblemCode.IO_ERROR,
            "disk failed",
            context,
            (List<GridGrindResponse.ProblemCause>) null);
    assertEquals(List.of(), explicitWithoutCauses.causes());

    GridGrindResponse.Problem explicit =
        GridGrindProblems.problem(
            GridGrindProblemCode.IO_ERROR, "disk failed", context, new IOException("disk failed"));
    assertEquals(GridGrindProblemCode.IO_ERROR, explicit.code());
    assertEquals(1, explicit.causes().size());

    GridGrindResponse.ProblemCause supplemental =
        GridGrindProblems.supplementalCause(
            "EXECUTE_REQUEST",
            new IOException("close failed"),
            "Workbook close failed after the primary problem");
    assertEquals("EXECUTE_REQUEST", supplemental.stage());
    assertTrue(supplemental.message().contains("close failed"));
    GridGrindResponse.ProblemCause supplementalWithoutPrefix =
        GridGrindProblems.supplementalCause(
            "EXECUTE_REQUEST", new IOException("close failed"), " ");
    assertEquals("close failed", supplementalWithoutPrefix.message());
    GridGrindResponse.ProblemCause supplementalWithNullPrefix =
        GridGrindProblems.supplementalCause(
            "EXECUTE_REQUEST", new IOException("close failed"), null);
    assertEquals("close failed", supplementalWithNullPrefix.message());

    GridGrindResponse.Problem augmented = GridGrindProblems.appendCause(explicit, supplemental);
    assertEquals(2, augmented.causes().size());

    GridGrindResponse.ProblemCause synthetic = GridGrindProblems.problemCause(explicit);
    assertEquals("GridGrindProblem", synthetic.type());
    assertEquals(GridGrindProblemCode.IO_ERROR.name(), synthetic.className());
    assertEquals("WRITE_RESPONSE", synthetic.stage());

    assertEquals(List.of(), GridGrindProblems.causesFor(null));
    assertEquals(
        "UnsupportedOperationException",
        GridGrindProblems.messageFor(new UnsupportedOperationException(" ")));
    assertEquals(
        "dev.erst.gridgrind.protocol.GridGrindProblemsTest$1",
        GridGrindProblems.messageFor(new RuntimeException(" ") {}));

    RuntimeException first = new RuntimeException("first");
    RuntimeException second = new RuntimeException("second");
    first.initCause(second);
    second.initCause(first);
    assertEquals(2, GridGrindProblems.causesFor(first).size());
  }

  @Test
  void enrichesPayloadContextFromPayloadExceptions() {
    // ReadRequest context is enriched with json location from PayloadException
    GridGrindResponse.ProblemContext readContext =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);

    GridGrindResponse.Problem problem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request",
                "analysis.sheets[0].previewRowCount",
                6,
                41,
                new IllegalArgumentException("bad")),
            readContext);

    assertEquals("analysis.sheets[0].previewRowCount", problem.context().jsonPath());
    assertEquals(6, problem.context().jsonLine());
    assertEquals(41, problem.context().jsonColumn());

    // Explicit json location in ReadRequest is preserved even when exception carries different
    // values
    GridGrindResponse.ProblemContext explicitContext =
        new GridGrindResponse.ProblemContext.ReadRequest(
            "/tmp/request.json", "analysis.sheets[1]", 9, 12);

    GridGrindResponse.Problem preserved =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request",
                "analysis.sheets[0].previewRowCount",
                6,
                41,
                new IllegalArgumentException("bad")),
            explicitContext);

    assertEquals("analysis.sheets[1]", preserved.context().jsonPath());
    assertEquals(9, preserved.context().jsonLine());
    assertEquals(12, preserved.context().jsonColumn());

    // ValidateRequest context is not enriched with json location (no json fields in that record)
    GridGrindResponse.ProblemContext validateContext =
        new GridGrindResponse.ProblemContext.ValidateRequest(null, null);
    GridGrindResponse.Problem notEnriched =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request",
                "analysis.sheets[0].previewRowCount",
                6,
                41,
                new IllegalArgumentException("bad")),
            validateContext);
    assertNull(notEnriched.context().jsonPath());
  }

  @Test
  void enrichesAnalyzeWorkbookContextFromFormulaException() {
    GridGrindResponse.ProblemContext analyzeContext =
        new GridGrindResponse.ProblemContext.AnalyzeWorkbook("NEW", "NONE", null, null, null);
    InvalidFormulaException formulaException =
        new InvalidFormulaException(
            "Budget", "C3", "SUM(", "bad formula", new IllegalArgumentException("bad"));

    GridGrindResponse.Problem enriched =
        GridGrindProblems.fromException(formulaException, analyzeContext);

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, enriched.code());
    assertEquals("ANALYZE_WORKBOOK", enriched.context().stage());
    assertEquals("Budget", enriched.context().sheetName());
    assertEquals("C3", enriched.context().address());
    assertEquals("SUM(", enriched.context().formula());
  }

  @Test
  void leavesReadRequestContextUnchangedWhenExceptionIsNotPayloadException() {
    GridGrindResponse.ProblemContext readContext =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);

    GridGrindResponse.Problem problem =
        GridGrindProblems.fromException(new IOException("disk"), readContext);

    assertEquals(GridGrindProblemCode.IO_ERROR, problem.code());
    assertEquals("READ_REQUEST", problem.context().stage());
    assertEquals("/tmp/request.json", problem.context().requestPath());
    assertNull(problem.context().jsonPath());
    assertNull(problem.context().jsonLine());
    assertNull(problem.context().jsonColumn());
  }

  @Test
  void enrichContextReturnsCatchAllContextForUnenrichedTypes() {
    // Contexts with no when-guard enrichment (ParseArguments, ValidateRequest, WriteResponse)
    // and the guard-false paths for ReadRequest and AnalyzeWorkbook always return context
    // unchanged.
    RuntimeException ex = new RuntimeException("test");

    GridGrindResponse.ProblemContext parseArgs =
        new GridGrindResponse.ProblemContext.ParseArguments("--request");
    assertSame(parseArgs, GridGrindProblems.enrichContext(parseArgs, ex));

    GridGrindResponse.ProblemContext validateReq =
        new GridGrindResponse.ProblemContext.ValidateRequest(null, null);
    assertSame(validateReq, GridGrindProblems.enrichContext(validateReq, ex));

    GridGrindResponse.ProblemContext writeResp =
        new GridGrindResponse.ProblemContext.WriteResponse("/tmp/response.json");
    assertSame(writeResp, GridGrindProblems.enrichContext(writeResp, ex));
  }

  @Test
  void enrichesFormulaAndRangeContextWithoutOverwritingExplicitValues() {
    GridGrindResponse.ProblemContext blankContext =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            null, null, null, null, null, null, null, null);
    InvalidFormulaException formulaException =
        new InvalidFormulaException(
            "Budget", "B4", "SUM(", "bad formula", new IllegalArgumentException("bad"));

    GridGrindResponse.Problem enriched =
        GridGrindProblems.fromException(formulaException, blankContext);

    assertEquals("Budget", enriched.context().sheetName());
    assertEquals("B4", enriched.context().address());
    assertEquals("SUM(", enriched.context().formula());

    GridGrindResponse.Problem rangeProblem =
        GridGrindProblems.fromException(
            new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad")),
            blankContext);
    assertEquals("A1:", rangeProblem.context().range());
    assertEquals(
        "A1:",
        GridGrindProblems.rangeFor(
            new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad"))));

    GridGrindResponse.ProblemContext explicitContext =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            null, null, null, null, "Summary", "C9", "C1:C9", "AVERAGE(C1:C8)");

    GridGrindResponse.Problem preserved =
        GridGrindProblems.fromException(formulaException, explicitContext);
    GridGrindResponse.Problem preservedRange =
        GridGrindProblems.fromException(
            new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad")),
            explicitContext);

    assertEquals("Summary", preserved.context().sheetName());
    assertEquals("C9", preserved.context().address());
    assertEquals("AVERAGE(C1:C8)", preserved.context().formula());
    assertEquals("C1:C9", preservedRange.context().range());
  }
}
