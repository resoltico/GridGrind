package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import org.junit.jupiter.api.Test;

/** Tests for the transport-neutral GridGrind request executor port. */
class GridGrindRequestExecutorTest {
  @Test
  void requireNonNullReturnsExecutorAndRejectsNull() {
    GridGrindRequestExecutor executor =
        (request, bindings, sink) ->
            GridGrindResponse.failure(
                GridGrindProtocolVersion.current(),
                GridGrindProblems.problem(
                    GridGrindProblemCode.INVALID_REQUEST,
                    "boom",
                    new GridGrindResponse.ProblemContext.ValidateRequest(null, null),
                    (Throwable) null));

    assertSame(executor, GridGrindRequestExecutor.requireNonNull(executor));
    assertThrows(NullPointerException.class, () -> GridGrindRequestExecutor.requireNonNull(null));
  }
}
