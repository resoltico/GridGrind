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
          "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
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
