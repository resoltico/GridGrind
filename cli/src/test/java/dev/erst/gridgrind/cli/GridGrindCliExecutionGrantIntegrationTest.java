package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Black-box proof that CLI execution authority is separate from an authored workbook plan. */
class GridGrindCliExecutionGrantIntegrationTest {
  @TempDir Path root;

  @Test
  void executesOnlyWhenTheCommandSuppliesAGrantThatPermitsThePlanOperation()
      throws java.io.IOException {
    Path request = writeRequest();
    Path allowedGrant = writeGrant(List.of("ENSURE_SHEET"));
    ByteArrayOutputStream allowedOutput = new ByteArrayOutputStream();

    int allowedExitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", request.toString(), "--grant", allowedGrant.toString()},
                InputStream.nullInputStream(),
                allowedOutput);

    assertEquals(0, allowedExitCode);
    assertInstanceOf(
        WorkbookResult.Success.class,
        GridGrindJson.readWorkbookResult(allowedOutput.toByteArray()));

    Path deniedGrant = writeGrant(List.of());
    ByteArrayOutputStream deniedOutput = new ByteArrayOutputStream();
    int deniedExitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", request.toString(), "--grant", deniedGrant.toString()},
                InputStream.nullInputStream(),
                deniedOutput);

    WorkbookResult.Failure denied =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            GridGrindJson.readWorkbookResult(deniedOutput.toByteArray()));
    assertEquals(1, deniedExitCode);
    assertEquals(GridGrindProblemCode.AUTHORITY_DENIED, denied.problem().code());
  }

  private Path writeRequest() throws java.io.IOException {
    Path request = root.resolve("request.json");
    Files.writeString(
        request,
        """
        {
          "protocolVersion": "V3",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "ensure-sheet",
              "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
              "action": { "type": "ENSURE_SHEET" }
            }
          ]
        }
        """,
        StandardCharsets.UTF_8);
    return request;
  }

  private Path writeGrant(List<String> operationIds) throws java.io.IOException {
    Path grant = Files.createTempFile("gridgrind-cli-execution-grant-", ".json");
    CliExecutionGrantDocument document =
        new CliExecutionGrantDocument(
            List.of(),
            operationIds,
            new CliGrantTargetAuthority.WorkbookWide(),
            new CliGrantPublicationAuthority.None(),
            List.of(),
            new CliGrantAcceptancePolicy.MinimumOnly());
    Files.write(grant, GridGrindCliJson.writeBytes(document));
    return grant;
  }
}
