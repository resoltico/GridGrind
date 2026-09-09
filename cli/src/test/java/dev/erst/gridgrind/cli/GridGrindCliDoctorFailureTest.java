package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Verifies that doctor-analysis runtime failures remain structured CLI reports. */
class GridGrindCliDoctorFailureTest extends GridGrindCliTestSupport {
  @Test
  void doctorRuntimeFailureUsesTheRequestRedactorAndReturnsAnInvalidReport() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-runtime-failure-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    try {
      GridGrindCli cli =
          new GridGrindCli(
              (request, inputs, sink) -> {
                throw new AssertionError("doctor requests must not execute");
              },
              new FailingDoctor(),
              new CliRequestReader(),
              new CliResponseWriter(),
              () -> false);

      int exitCode =
          cli.run(
              new String[] {"--doctor-request", "--execution-root", workspace.toString()},
              new ByteArrayInputStream(
                  minimalRequestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                      .getBytes(StandardCharsets.UTF_8)),
              stdout,
              ByteArrayOutputStream.nullOutputStream());

      RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());
      assertEquals(1, exitCode);
      assertFalse(report.valid());
    } finally {
      CliExecutionBindingsFactory.deleteTreeIfExists(workspace);
    }
  }

  /** Doctor double that demonstrates containment of unexpected analysis failures. */
  private static final class FailingDoctor implements GridGrindRequestDoctor {
    @Override
    public RequestDoctorReport diagnose(RequestAnalysis analysis, RequestInput requestInput) {
      throw failure();
    }

    @Override
    public RequestDoctorReport diagnose(
        RequestAnalysis analysis, RequestInput requestInput, GridGrindRequestInputs inputs) {
      throw failure();
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request) {
      throw failure();
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs) {
      throw failure();
    }

    private static IllegalStateException failure() {
      return new IllegalStateException("deliberate doctor failure");
    }
  }
}
