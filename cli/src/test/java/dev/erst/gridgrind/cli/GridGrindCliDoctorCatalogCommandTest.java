package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;
import org.junit.jupiter.api.Test;

/** Focused integration coverage for doctor-request catalog command flows. */
class GridGrindCliDoctorCatalogCommandTest extends GridGrindCliTestSupport {
  @Test
  void doctorRequestReturnsStructuredInvalidReportForSemanticallyInvalidRequests()
      throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-invalid-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"OVERWRITE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
    assertEquals("NEW", report.summary().orElseThrow().sourceType());
  }

  @Test
  void doctorRequestClassifiesMissingRequiredRootFieldsAsShapeFailures() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-doctor-missing-root-", ".json");
    Files.writeString(
        requestPath,
        """
        {
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "execution": {
            "mode": { "type": "FULL_XSSF" },
            "journal": { "level": "SUMMARY" },
            "calculation": {
              "strategy": { "type": "DO_NOT_CALCULATE" },
              "markRecalculateOnOpen": false
            }
          },
          "formulaEnvironment": {
            "externalWorkbooks": [],
            "missingWorkbookPolicy": "ERROR",
            "udfToolpacks": []
          },
          "steps": []
        }
        """);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST_SHAPE, report.primaryProblem().orElseThrow().code());
    assertEquals(java.util.Optional.of("protocolVersion"), readRequestContext(report).jsonPath());
  }

  @Test
  void doctorRequestWritesProblemPointerToStderrWhenResponsePathCapturesInvalidReport()
      throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-invalid-response-");
    Path responsePath = Files.createTempFile("gridgrind-invalid-doctor-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--doctor-request",
                  "--execution-root",
                  workspace.toString(),
                  "--response",
                  responsePath.toString()
                },
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"OVERWRITE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report =
        GridGrindJson.readRequestDoctorReport(Files.readAllBytes(responsePath));

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(
        "GridGrind wrote the doctor report to "
            + responsePath.toAbsolutePath()
            + "; inspect that file for problems [INVALID_REQUEST: OVERWRITE persistence requires an"
            + " EXISTING source; a NEW workbook has no source file to overwrite]."
            + System.lineSeparator(),
        stderr.toString(StandardCharsets.UTF_8));
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
  }

  @Test
  void doctorRequestPreflightsExistingWorkbookSourcesFromTheRequestFileDirectory()
      throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-doctor-source-");
    Path requestPath = requestDirectory.resolve("doctor-existing.json");
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"EXISTING\", \"path\": \"missing-workbook.xlsx\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              {
                "stepId": "summary",
                "target": { "type": "WORKBOOK_CURRENT" },
                "query": { "type": "GET_WORKBOOK_SUMMARY" }
              }
            ]
            """));
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND, report.primaryProblem().orElseThrow().code());
    assertEquals("OPEN_WORKBOOK", report.primaryProblem().orElseThrow().context().stage());
    assertEquals(
        java.util.Optional.of(requestDirectory.resolve("missing-workbook.xlsx").toString()),
        openWorkbookContext(report).sourceWorkbookPath());
  }

  @Test
  void doctorRequestReturnsCompactReadFailureWhenTheRequestFileCannotBeOpened() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    Path missingRequestPath =
        Path.of("tmp", "doctor-missing-" + UUID.randomUUID() + ".json").toAbsolutePath();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", missingRequestPath.toString()},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.code());
    assertEquals("doctor-request", failure.command());
    assertEquals(java.util.Optional.of("--request"), failure.argument());
    assertEquals("Request file not found: " + missingRequestPath, failure.message());
  }

  @Test
  void doctorRequestRequiresExecutionRootWhenRequestArrivesOnStdin() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request"},
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals("doctor-request", failure.command());
    assertEquals(java.util.Optional.of("--execution-root"), failure.argument());
    assertTrue(failure.message().contains("--execution-root"));
  }
}
