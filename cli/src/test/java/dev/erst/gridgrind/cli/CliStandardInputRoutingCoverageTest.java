package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
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

    WorkbookResult.Success response =
        assertInstanceOf(
            WorkbookResult.Success.class, GridGrindJson.readWorkbookResult(stdout.toByteArray()));

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

    CommandError failure = commandErrorOnStdout(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.primaryProblem().code());
    assertEquals("--request", parseArgumentsContext(failure).argumentName().orElseThrow());
    assertEquals(
        GridGrindRequestSurfaceContractText.standardInputRequiresRequestMessage(),
        failure.primaryProblem().message());
  }
}
