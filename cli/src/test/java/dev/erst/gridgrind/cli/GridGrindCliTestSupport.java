package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Objects;

/** Shared helpers for CLI integration tests. */
@SuppressWarnings("PMD.UseUtilityClass")
class GridGrindCliTestSupport {
  protected GridGrindCliTestSupport() {}

  protected static String requestJson(String sourceJson, String persistenceJson, String stepsJson) {
    return requestJson(
        sourceJson,
        persistenceJson,
        defaultExecutionJson(),
        emptyFormulaEnvironmentJson(),
        stepsJson);
  }

  protected static String requestJson(
      String sourceJson,
      String persistenceJson,
      String executionJson,
      String formulaEnvironmentJson,
      String stepsJson) {
    return """
        {
          "protocolVersion": "V1",
          "source": %s,
          "persistence": %s,
          "execution": %s,
          "formulaEnvironment": %s,
          "steps": %s
        }
        """
        .formatted(sourceJson, persistenceJson, executionJson, formulaEnvironmentJson, stepsJson);
  }

  protected static String requestJsonWithPlanId(
      String planId,
      String sourceJson,
      String persistenceJson,
      String executionJson,
      String formulaEnvironmentJson,
      String stepsJson) {
    return """
        {
          "protocolVersion": "V1",
          "planId": "%s",
          "source": %s,
          "persistence": %s,
          "execution": %s,
          "formulaEnvironment": %s,
          "steps": %s
        }
        """
        .formatted(
            planId, sourceJson, persistenceJson, executionJson, formulaEnvironmentJson, stepsJson);
  }

  protected static String defaultExecutionJson() {
    return executionJson("NORMAL", "DO_NOT_CALCULATE", false);
  }

  protected static String verboseExecutionJson() {
    return executionJson("VERBOSE", "DO_NOT_CALCULATE", false);
  }

  protected static String evaluateAllExecutionJson() {
    return executionJson("NORMAL", "EVALUATE_ALL", false);
  }

  protected static String executionJson(
      String journalLevel, String calculationStrategy, boolean markRecalculateOnOpen) {
    return """
        {
          "mode": {"type": "FULL_XSSF"},
          "journal": { "level": "%s" },
          "calculation": {
            "strategy": { "type": "%s" },
            "markRecalculateOnOpen": %s
          }
        }
        """
        .formatted(journalLevel, calculationStrategy, markRecalculateOnOpen);
  }

  protected static String emptyFormulaEnvironmentJson() {
    return """
        {
          "externalWorkbooks": [],
          "missingWorkbookPolicy": "ERROR",
          "udfToolpacks": []
        }
        """;
  }

  protected static ProblemContext.ParseArguments parseArgumentsContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.ParseArguments.class, failure.problem().context());
  }

  protected static ProblemContext.ParseArguments parseArgumentsContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.ParseArguments.class, report.primaryProblem().orElseThrow().context());
  }

  protected static ProblemContext.ReadRequest readRequestContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.ReadRequest.class, failure.problem().context());
  }

  protected static ProblemContext.ReadRequest readRequestContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.ReadRequest.class, report.primaryProblem().orElseThrow().context());
  }

  protected static ProblemContext.ResolveInputs resolveInputsContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.ResolveInputs.class, report.primaryProblem().orElseThrow().context());
  }

  protected static ProblemContext.OpenWorkbook openWorkbookContext(RequestDoctorReport report) {
    return assertInstanceOf(
        ProblemContext.OpenWorkbook.class, report.primaryProblem().orElseThrow().context());
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
        ProblemContext.WriteResponse.class, report.primaryProblem().orElseThrow().context());
  }

  protected static ProblemContext.ExecuteStep executeStepContext(
      GridGrindResponse.Failure failure) {
    return assertInstanceOf(ProblemContext.ExecuteStep.class, failure.problem().context());
  }

  protected static CliFailureReport cliFailure(byte[] bytes) throws IOException {
    return GridGrindCliJson.readBytes(bytes, CliFailureReport.class);
  }

  /**
   * Reads a {@link CliFailureReport} from stderr and asserts that stdout stayed empty.
   *
   * <p>Use this helper in every test that expects a CLI failure with no {@code --response} path
   * configured. It encodes the current routing contract — structured JSON on stderr, nothing on
   * stdout — in a single call so individual tests cannot forget either half.
   */
  protected static CliFailureReport cliFailureOnStderr(
      ByteArrayOutputStream stdout, ByteArrayOutputStream stderr) throws IOException {
    assertEquals(
        "",
        stdout.toString(StandardCharsets.UTF_8),
        "stdout must be empty when CLI failure is routed to stderr");
    return GridGrindCliJson.readBytes(stderr.toByteArray(), CliFailureReport.class);
  }

  protected static GridGrindResponse response(
      ByteArrayOutputStream stdout, ByteArrayOutputStream stderr) throws IOException {
    return GridGrindJson.readResponse(nonEmptyPayload(stdout, stderr));
  }

  protected static RequestDoctorReport doctorReport(
      ByteArrayOutputStream stdout, ByteArrayOutputStream stderr) throws IOException {
    return GridGrindJson.readRequestDoctorReport(nonEmptyPayload(stdout, stderr));
  }

  private static byte[] nonEmptyPayload(
      ByteArrayOutputStream primary, ByteArrayOutputStream secondary) {
    Objects.requireNonNull(primary, "primary must not be null");
    Objects.requireNonNull(secondary, "secondary must not be null");
    byte[] primaryBytes = primary.toByteArray();
    if (primaryBytes.length > 0) {
      return primaryBytes;
    }
    return secondary.toByteArray();
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
