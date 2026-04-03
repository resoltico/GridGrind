package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.protocol.dto.*;
import org.junit.jupiter.api.Test;

/** Tests for the transport-neutral GridGrind request executor port. */
class GridGrindRequestExecutorTest {
  @Test
  void requireNonNullReturnsExecutorAndRejectsNull() {
    GridGrindRequestExecutor executor =
        request ->
            new GridGrindResponse.Failure(
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
