package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResults;
import org.junit.jupiter.api.Test;

/** Tests for the transport-neutral GridGrind request executor port. */
class GridGrindRequestExecutorTest {
  @Test
  void requireNonNullReturnsExecutorAndRejectsNull() {
    GridGrindRequestExecutor executor =
        (request, bindings, sink) ->
            WorkbookResults.failure(
                GridGrindProtocolVersion.current(),
                GridGrindProblems.problem(
                    GridGrindProblemCode.INVALID_REQUEST,
                    "boom",
                    new dev.erst.gridgrind.core.problem.ProblemContext.ValidateRequest(
                        dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape
                            .unknown()),
                    (Throwable) null));

    assertSame(executor, GridGrindRequestExecutor.requireNonNull(executor));
    assertThrows(NullPointerException.class, () -> GridGrindRequestExecutor.requireNonNull(null));
  }
}
