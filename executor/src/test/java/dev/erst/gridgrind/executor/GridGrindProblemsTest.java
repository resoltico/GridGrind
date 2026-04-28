package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCategory;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidSigningConfigurationException;
import dev.erst.gridgrind.excel.InvalidWorkbookPasswordException;
import dev.erst.gridgrind.excel.WorkbookPasswordRequiredException;
import dev.erst.gridgrind.excel.WorkbookSecurityException;
import java.io.IOException;
import java.time.DateTimeException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for GridGrindProblems exception classification and context enrichment. */
class GridGrindProblemsTest {
  @Test
  void classifiesProblemCodesAcrossProtocolAndTransportFamilies() {
    assertEquals(
        GridGrindProblemCode.INVALID_JSON,
        GridGrindProblems.codeFor(
            new InvalidJsonException(
                "bad json", "steps[0]", 4, 12, new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST_SHAPE,
        GridGrindProblems.codeFor(
            new InvalidRequestShapeException(
                "bad shape", "steps[0]", 4, 12, new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(
            new InvalidRequestException(
                "bad request",
                "steps[0].target.rowCount",
                6,
                41,
                new IllegalArgumentException("bad"))));
    assertEquals(GridGrindProblemCode.IO_ERROR, GridGrindProblems.codeFor(new IOException("disk")));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(new IllegalArgumentException("bad")));
    assertEquals(
        GridGrindProblemCode.WORKBOOK_PASSWORD_REQUIRED,
        GridGrindProblems.codeFor(
            new WorkbookPasswordRequiredException(java.nio.file.Path.of("/tmp/encrypted.xlsx"))));
    assertEquals(
        GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD,
        GridGrindProblems.codeFor(
            new InvalidWorkbookPasswordException(java.nio.file.Path.of("/tmp/encrypted.xlsx"))));
    assertEquals(
        GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION,
        GridGrindProblems.codeFor(new InvalidSigningConfigurationException("bad signing")));
    assertEquals(
        GridGrindProblemCode.WORKBOOK_SECURITY_ERROR,
        GridGrindProblems.codeFor(new WorkbookSecurityException("crypto failed", null)));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindProblems.codeFor(new DateTimeException("bad date")));
    AssertionFailure assertionFailure =
        new AssertionFailure(
            "assert-total",
            "EXPECT_CELL_VALUE",
            new CellSelector.ByAddress("Budget", "B4"),
            new Assertion.CellValue(new ExpectedCellValue.NumericValue(42.0d)),
            List.of());
    assertEquals(
        GridGrindProblemCode.ASSERTION_FAILED,
        GridGrindProblems.codeFor(
            new AssertionFailedException("assertion failed", assertionFailure)));
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        GridGrindProblems.codeFor(new UnsupportedOperationException("boom")));
  }

  @Test
  void buildsStructuredProblemsAndPublicDiagnostics() {
    ProblemContext context =
        new ProblemContext.ValidateRequest(ProblemContext.RequestShape.known("NEW", "NONE"));
    InvalidRequestException exception =
        new InvalidRequestException(
            "bad request",
            "steps[0].target.rowCount",
            6,
            41,
            new IllegalArgumentException("root cause"));
    GridGrindResponse.Problem problem = GridGrindProblems.fromException(exception, context);

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, problem.code());
    assertEquals(GridGrindProblemCategory.REQUEST, problem.category());
    assertEquals(1, problem.causes().size());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, problem.causes().getFirst().code());
    assertEquals("VALIDATE_REQUEST", problem.causes().getFirst().stage());
    assertSame(context, problem.context());
  }

  @Test
  void supportsExplicitProblemsAndSupplementalDiagnostics() {
    ProblemContext context =
        new ProblemContext.WriteResponse(
            ProblemContext.ResponseOutput.responseFile("/tmp/response.json"));

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
    assertEquals(GridGrindProblemCode.IO_ERROR, GridGrindProblems.problemCause(explicit).code());

    GridGrindResponse.ProblemCause withoutPrefix =
        GridGrindProblems.supplementalCause("EXECUTE_REQUEST", new IOException("disk failed"), "");
    assertEquals("disk failed", withoutPrefix.message());
    assertEquals(GridGrindProblemCode.IO_ERROR, withoutPrefix.code());
    assertEquals(
        "disk failed",
        GridGrindProblems.supplementalCause("EXECUTE_REQUEST", new IOException("disk failed"), null)
            .message());

    GridGrindResponse.Problem withoutCauses =
        GridGrindProblems.problem(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new ProblemContext.ValidateRequest(ProblemContext.RequestShape.known("NEW", "NONE")),
            (List<GridGrindResponse.ProblemCause>) null);
    assertEquals(List.of(), withoutCauses.causes());

    GridGrindResponse.Problem assertionProblem =
        GridGrindProblems.fromException(
            new AssertionFailedException(
                "assertion failed",
                new AssertionFailure(
                    "assert-total",
                    "EXPECT_CELL_VALUE",
                    new CellSelector.ByAddress("Budget", "B4"),
                    new Assertion.CellValue(new ExpectedCellValue.NumericValue(42.0d)),
                    List.of())),
            new ProblemContext.ExecuteStep(
                ProblemContext.RequestShape.known("NEW", "NONE"),
                new ProblemContext.StepReference(
                    1, "assert-total", "ASSERTION", "EXPECT_CELL_VALUE"),
                ProblemContext.ProblemLocation.cell("Budget", "B4")));
    assertEquals(GridGrindProblemCategory.ASSERTION, assertionProblem.category());
    assertTrue(assertionProblem.assertionFailure().isPresent());
  }

  @Test
  void enrichesPayloadContextFromPayloadExceptions() {
    ProblemContext.ReadRequest readContext =
        new ProblemContext.ReadRequest(
            ProblemContext.RequestInput.requestFile("/tmp/request.json"),
            ProblemContext.JsonLocation.unavailable());

    GridGrindResponse.Problem problem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request",
                "steps[0].target.rowCount",
                6,
                41,
                new IllegalArgumentException("bad")),
            readContext);

    assertEquals(
        java.util.Optional.of("steps[0].target.rowCount"),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(problem).jsonPath());
    assertEquals(
        java.util.Optional.of(6),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(problem).jsonLine());
    assertEquals(
        java.util.Optional.of(41),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(problem).jsonColumn());

    GridGrindResponse.Problem lineColumnProblem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request", null, 9, 4, new IllegalArgumentException("bad")),
            readContext);
    GridGrindResponse.Problem unavailableProblem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request", "steps[0]", null, null, new IllegalArgumentException("bad")),
            readContext);
    GridGrindResponse.Problem partiallyUnavailableProblem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request", "steps[0]", 11, null, new IllegalArgumentException("bad")),
            readContext);

    assertEquals(
        java.util.Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(lineColumnProblem)
            .jsonPath());
    assertEquals(
        java.util.Optional.of(9),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(lineColumnProblem)
            .jsonLine());
    assertEquals(
        java.util.Optional.of(4),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(lineColumnProblem)
            .jsonColumn());
    assertEquals(
        java.util.Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(unavailableProblem)
            .jsonLine());
    assertEquals(
        java.util.Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(unavailableProblem)
            .jsonColumn());
    assertEquals(
        java.util.Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(partiallyUnavailableProblem)
            .jsonLine());
  }

  @Test
  void enrichesExecuteCalculationContextFromFormulaExceptions() {
    GridGrindResponse.Problem problem =
        GridGrindProblems.fromException(
            new InvalidFormulaException("Budget", "B1", "SUM(", "bad formula", null),
            new ProblemContext.ExecuteCalculation.Preflight(
                ProblemContext.RequestShape.known("NEW", "SAVE_AS"),
                ProblemContext.ProblemLocation.unknown()));

    assertEquals("CALCULATION_PREFLIGHT", problem.context().stage());
    assertEquals(
        java.util.Optional.of("Budget"),
        DefaultGridGrindRequestExecutorTestSupport.calculationPreflightContext(problem)
            .sheetName());
    assertEquals(
        java.util.Optional.of("B1"),
        DefaultGridGrindRequestExecutorTestSupport.calculationPreflightContext(problem).address());
    assertEquals(
        java.util.Optional.of("SUM("),
        DefaultGridGrindRequestExecutorTestSupport.calculationPreflightContext(problem).formula());
  }

  @Test
  void leavesAllPassthroughContextsUnchanged() {
    RuntimeException exception = new RuntimeException("test");

    ProblemContext.ParseArguments parseArguments =
        new ProblemContext.ParseArguments(ProblemContext.CliArgument.named("--request"));
    ProblemContext.ValidateRequest validateRequest =
        new ProblemContext.ValidateRequest(ProblemContext.RequestShape.unknown());
    ProblemContext.OpenWorkbook openWorkbook =
        new ProblemContext.OpenWorkbook(
            ProblemContext.RequestShape.known("EXISTING", "NONE"),
            ProblemContext.WorkbookReference.existingFile("/tmp/source.xlsx"));
    ProblemContext.PersistWorkbook persistWorkbook =
        new ProblemContext.PersistWorkbook(
            ProblemContext.RequestShape.known("NEW", "SAVE_AS"),
            ProblemContext.PersistenceReference.saveAs("/tmp/output.xlsx"));
    ProblemContext.ExecuteRequest executeRequest =
        new ProblemContext.ExecuteRequest(ProblemContext.RequestShape.known("NEW", "NONE"));
    ProblemContext.WriteResponse writeResponse =
        new ProblemContext.WriteResponse(
            ProblemContext.ResponseOutput.responseFile("/tmp/response.json"));
    ProblemContext.ExecuteStep executeStep =
        new ProblemContext.ExecuteStep(
            ProblemContext.RequestShape.unknown(),
            new ProblemContext.StepReference(0, "step-0", "MUTATION", "SET_CELL"),
            ProblemContext.ProblemLocation.unknown());

    assertSame(parseArguments, GridGrindProblems.enrichContext(parseArguments, exception));
    assertSame(validateRequest, GridGrindProblems.enrichContext(validateRequest, exception));
    assertSame(openWorkbook, GridGrindProblems.enrichContext(openWorkbook, exception));
    assertSame(persistWorkbook, GridGrindProblems.enrichContext(persistWorkbook, exception));
    assertSame(executeRequest, GridGrindProblems.enrichContext(executeRequest, exception));
    assertSame(writeResponse, GridGrindProblems.enrichContext(writeResponse, exception));
    assertEquals(executeStep, GridGrindProblems.enrichContext(executeStep, exception));

    ProblemContext.ExecuteCalculation calculation =
        new ProblemContext.ExecuteCalculation.Preflight(
            ProblemContext.RequestShape.known("NEW", "NONE"),
            ProblemContext.ProblemLocation.unknown());
    assertSame(
        calculation, GridGrindProblems.enrichContext(calculation, new AssertionError("boom")));
    assertSame(
        executeStep, GridGrindProblems.enrichContext(executeStep, new AssertionError("boom")));
  }

  @Test
  void fallsBackToAnonymousTypeNameAndReturnsSinglePublicCause() {
    Throwable anonymous =
        new Throwable(" ") {
          @Override
          public Throwable getCause() {
            return this;
          }
        };

    assertTrue(GridGrindProblems.messageFor(anonymous).contains("$"));
    assertEquals(1, GridGrindProblems.causesFor(anonymous).size());
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        GridGrindProblems.causesFor(anonymous).getFirst().code());

    ProblemContext.ReadRequest readContext =
        new ProblemContext.ReadRequest(
            ProblemContext.RequestInput.requestFile("/tmp/request.json"),
            ProblemContext.JsonLocation.unavailable());
    assertSame(readContext, GridGrindProblems.enrichContext(readContext, anonymous));
  }
}
