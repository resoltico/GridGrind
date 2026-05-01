package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import org.junit.jupiter.api.Test;

/** Tests for the transport-neutral GridGrind request executor port. */
class GridGrindRequestExecutorTest {
  @Test
  void requireNonNullReturnsExecutorAndRejectsNull() {
    GridGrindRequestExecutor executor =
        (request, bindings, sink) ->
            GridGrindResponses.failure(
                GridGrindProtocolVersion.current(),
                GridGrindProblems.problem(
                    GridGrindProblemCode.INVALID_REQUEST,
                    "boom",
                    new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                        dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                            .unknown()),
                    (Throwable) null));

    assertSame(executor, GridGrindRequestExecutor.requireNonNull(executor));
    assertThrows(NullPointerException.class, () -> GridGrindRequestExecutor.requireNonNull(null));
  }
}
