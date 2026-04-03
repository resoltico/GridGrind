package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.protocol.dto.GridGrindProblemCategory;
import dev.erst.gridgrind.protocol.dto.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.json.InvalidJsonException;
import dev.erst.gridgrind.protocol.json.InvalidRequestException;
import dev.erst.gridgrind.protocol.json.InvalidRequestShapeException;
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
  void buildsStructuredProblemsAndPublicDiagnostics() {
    GridGrindResponse.ProblemContext context =
        new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE");
    InvalidRequestException exception =
        new InvalidRequestException(
            "bad request", "reads[0].rowCount", 6, 41, new IllegalArgumentException("root cause"));
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
    GridGrindResponse.ProblemContext.ApplyOperation applyOperation =
        new GridGrindResponse.ProblemContext.ApplyOperation(
            null, null, null, null, null, null, null, null, null);
    GridGrindResponse.ProblemContext.ExecuteRead executeRead =
        new GridGrindResponse.ProblemContext.ExecuteRead(
            null, null, null, null, null, null, null, null, null);

    assertSame(parseArguments, GridGrindProblems.enrichContext(parseArguments, exception));
    assertSame(validateRequest, GridGrindProblems.enrichContext(validateRequest, exception));
    assertSame(openWorkbook, GridGrindProblems.enrichContext(openWorkbook, exception));
    assertSame(persistWorkbook, GridGrindProblems.enrichContext(persistWorkbook, exception));
    assertSame(executeRequest, GridGrindProblems.enrichContext(executeRequest, exception));
    assertSame(writeResponse, GridGrindProblems.enrichContext(writeResponse, exception));
    assertSame(applyOperation, GridGrindProblems.enrichContext(applyOperation, exception));
    assertSame(executeRead, GridGrindProblems.enrichContext(executeRead, exception));
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

    GridGrindResponse.ProblemContext.ReadRequest readContext =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", null, null, null);
    assertSame(readContext, GridGrindProblems.enrichContext(readContext, anonymous));
  }
}
