package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.time.DateTimeException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for GridGrindProblems exception classification and context enrichment. */
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
        GridGrindProblemCode.NAMED_RANGE_NOT_FOUND,
        GridGrindProblems.codeFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
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
                "bad json", "reads[0]", 4, 12, new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST_SHAPE,
        GridGrindProblems.codeFor(
            new InvalidRequestShapeException(
                "bad shape", "reads[0]", 4, 12, new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(
            new InvalidRequestException(
                "bad request", "reads[0].rowCount", 6, 41, new IllegalArgumentException("bad"))));
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
            null, null, null, null, "Budget", "B4", null, "SUM(", null);

    InvalidFormulaException exception =
        new InvalidFormulaException(
            "Budget", "B4", "SUM(", "bad formula", new IllegalArgumentException("root cause"));
    GridGrindResponse.Problem problem = GridGrindProblems.fromException(exception, context);

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, problem.code());
    assertEquals(GridGrindProblemCategory.FORMULA, problem.category());
    assertEquals(2, problem.causes().size());
    assertEquals("SUM(", GridGrindProblems.formulaFor(exception));
    assertEquals("Budget", GridGrindProblems.sheetNameFor(exception));
    assertEquals("B4", GridGrindProblems.addressFor(exception));
    assertEquals("A1", GridGrindProblems.addressFor(new CellNotFoundException("A1")));
    assertEquals(
        "BAD!",
        GridGrindProblems.addressFor(
            new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))));
    assertNull(GridGrindProblems.rangeFor(new IllegalArgumentException("bad")));
    assertNull(GridGrindProblems.namedRangeNameFor(new IllegalArgumentException("bad")));
  }

  @Test
  void supportsExplicitProblemsAndSupplementalDiagnostics() {
    GridGrindResponse.ProblemContext context =
        new GridGrindResponse.ProblemContext.WriteResponse("/tmp/response.json");

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

    GridGrindResponse.Problem augmented = GridGrindProblems.appendCause(explicit, supplemental);
    assertEquals(2, augmented.causes().size());
    assertEquals("GridGrindProblem", GridGrindProblems.problemCause(explicit).type());

    GridGrindResponse.ProblemCause withoutPrefix =
        GridGrindProblems.supplementalCause("EXECUTE_REQUEST", new IOException("disk failed"), "");
    assertEquals("disk failed", withoutPrefix.message());
    assertEquals(
        "disk failed",
        GridGrindProblems.supplementalCause("EXECUTE_REQUEST", new IOException("disk failed"), null)
            .message());

    GridGrindResponse.Problem withoutCauses =
        GridGrindProblems.problem(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"),
            (List<GridGrindResponse.ProblemCause>) null);
    assertEquals(List.of(), withoutCauses.causes());
  }

  @Test
  void enrichesPayloadContextFromPayloadExceptions() {
    GridGrindResponse.ProblemContext.ReadRequest readContext =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);

    GridGrindResponse.Problem problem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request", "reads[0].rowCount", 6, 41, new IllegalArgumentException("bad")),
            readContext);

    assertEquals("reads[0].rowCount", problem.context().jsonPath());
    assertEquals(6, problem.context().jsonLine());
    assertEquals(41, problem.context().jsonColumn());
  }

  @Test
  void enrichesExecuteReadContextFromFormulaAndNamedRangeExceptions() {
    GridGrindResponse.ProblemContext.ExecuteRead executeRead =
        new GridGrindResponse.ProblemContext.ExecuteRead(
            "NEW", "NONE", 0, "GET_CELLS", "cells", null, null, null, null);
    InvalidFormulaException formulaException =
        new InvalidFormulaException(
            "Budget", "C3", "SUM(", "bad formula", new IllegalArgumentException("bad"));

    GridGrindResponse.Problem enrichedFormula =
        GridGrindProblems.fromException(formulaException, executeRead);

    assertEquals("Budget", enrichedFormula.context().sheetName());
    assertEquals("C3", enrichedFormula.context().address());
    assertEquals("SUM(", enrichedFormula.context().formula());

    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    GridGrindResponse.Problem enrichedNamedRange =
        GridGrindProblems.fromException(missingNamedRange, executeRead);

    assertEquals(GridGrindProblemCode.NAMED_RANGE_NOT_FOUND, enrichedNamedRange.code());
    assertEquals("BudgetTotal", enrichedNamedRange.context().namedRangeName());
  }

  @Test
  void enrichesApplyOperationContextWithoutOverwritingExplicitValues() {
    GridGrindResponse.ProblemContext.ApplyOperation blankContext =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            null, null, null, null, null, null, null, null, null);
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

    GridGrindResponse.Problem namedRangeProblem =
        GridGrindProblems.fromException(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()),
            blankContext);
    assertEquals("BudgetTotal", namedRangeProblem.context().namedRangeName());
  }

  @Test
  void leavesAllPassthroughContextsUnchanged() {
    RuntimeException exception = new RuntimeException("test");

    GridGrindResponse.ProblemContext.ParseArguments parseArguments =
        new GridGrindResponse.ProblemContext.ParseArguments("--request");
    GridGrindResponse.ProblemContext.ValidateRequest validateRequest =
        new GridGrindResponse.ProblemContext.ValidateRequest(null, null);
    GridGrindResponse.ProblemContext.OpenWorkbook openWorkbook =
        new GridGrindResponse.ProblemContext.OpenWorkbook("EXISTING", "NONE", "/tmp/source.xlsx");
    GridGrindResponse.ProblemContext.PersistWorkbook persistWorkbook =
        new GridGrindResponse.ProblemContext.PersistWorkbook(
            "NEW", "SAVE_AS", null, "/tmp/output.xlsx");
    GridGrindResponse.ProblemContext.ExecuteRequest executeRequest =
        new GridGrindResponse.ProblemContext.ExecuteRequest("NEW", "NONE");
    GridGrindResponse.ProblemContext.WriteResponse writeResponse =
        new GridGrindResponse.ProblemContext.WriteResponse("/tmp/response.json");

    assertSame(parseArguments, GridGrindProblems.enrichContext(parseArguments, exception));
    assertSame(validateRequest, GridGrindProblems.enrichContext(validateRequest, exception));
    assertSame(openWorkbook, GridGrindProblems.enrichContext(openWorkbook, exception));
    assertSame(persistWorkbook, GridGrindProblems.enrichContext(persistWorkbook, exception));
    assertSame(executeRequest, GridGrindProblems.enrichContext(executeRequest, exception));
    assertSame(writeResponse, GridGrindProblems.enrichContext(writeResponse, exception));
  }

  @Test
  void fallsBackToAnonymousTypeNameAndStopsCauseCycles() {
    Throwable anonymous =
        new Throwable(" ") {
          @Override
          public Throwable getCause() {
            return this;
          }
        };

    assertTrue(GridGrindProblems.messageFor(anonymous).contains("$"));
    assertEquals(1, GridGrindProblems.causesFor(anonymous).size());

    GridGrindResponse.ProblemContext.ReadRequest readContext =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);
    assertSame(readContext, GridGrindProblems.enrichContext(readContext, anonymous));
  }
}
