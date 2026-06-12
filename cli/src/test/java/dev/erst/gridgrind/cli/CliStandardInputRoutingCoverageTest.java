package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Covers request-path stdin sentinel routing through the live CLI transport. */
class CliStandardInputRoutingCoverageTest extends GridGrindCliTestSupport {
  @Test
  void requestDashReadsTheRequestFromStdinAndRootsExecutionAtExecutionRoot() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-request-dash-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", "-", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                            [
                              {
                                "stepId": "ensure-budget",
                                "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                                "action": { "type": "ENSURE_SHEET" }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse.Success response =
        assertInstanceOf(
            GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(0, exitCode);
    assertEquals(0, response.assertions().size());
    assertEquals(0, response.inspections().size());
  }

  @Test
  void requestDashRejectsRequestsThatAlsoBindStandardInputPayloads() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-request-dash-stdin-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", "-", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                            [
                              {
                                "stepId": "set-title",
                                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                                "action": {
                                  "type": "SET_CELL",
                                  "value": { "type": "TEXT", "source": { "type": "STANDARD_INPUT" } }
                                }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals("--request", failure.argument().orElseThrow());
    assertEquals(GridGrindContractText.standardInputRequiresRequestMessage(), failure.message());
  }
}
