package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import java.io.ByteArrayInputStream;
import java.io.IOException;

/** Shared helpers for CLI integration tests. */
@SuppressWarnings("PMD.UseUtilityClass")
class GridGrindCliTestSupport {
  protected GridGrindCliTestSupport() {}

  protected static ProblemContext.ParseArguments parseArgumentsContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.ParseArguments.class, failure.problem().context());
  }

  protected static ProblemContext.ParseArguments parseArgumentsContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.ParseArguments.class, report.problem().orElseThrow().context());
  }

  protected static ProblemContext.ReadRequest readRequestContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.ReadRequest.class, failure.problem().context());
  }

  protected static ProblemContext.ReadRequest readRequestContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.ReadRequest.class, report.problem().orElseThrow().context());
  }

  protected static ProblemContext.ResolveInputs resolveInputsContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.ResolveInputs.class, report.problem().orElseThrow().context());
  }

  protected static ProblemContext.OpenWorkbook openWorkbookContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.OpenWorkbook.class, report.problem().orElseThrow().context());
  }

  protected static ProblemContext.ExecuteRequest executeRequestContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.ExecuteRequest.class, failure.problem().context());
  }

  protected static ProblemContext.WriteResponse writeResponseContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.WriteResponse.class, failure.problem().context());
  }

  protected static ProblemContext.WriteResponse writeResponseContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.WriteResponse.class, report.problem().orElseThrow().context());
  }

  protected static ProblemContext.ExecuteStep executeStepContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.ExecuteStep.class, failure.problem().context());
  }

  /** ByteArrayInputStream that records whether {@code close()} was called. */
  protected static final class TrackingInputStream extends ByteArrayInputStream {
    private boolean closed;

    TrackingInputStream(byte[] bytes) {
      super(bytes);
    }

    @Override
    public void close() throws IOException {
      closed = true;
      super.close();
    }

    boolean closed() {
      return closed;
    }
  }
}
