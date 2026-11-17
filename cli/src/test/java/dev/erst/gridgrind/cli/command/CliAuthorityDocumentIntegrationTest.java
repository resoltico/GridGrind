package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.result.CommandError;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Authority refusals retain their input role across both commands and primary transports. */
class CliAuthorityDocumentIntegrationTest {
  @TempDir Path root;

  @ParameterizedTest
  @CsvSource({"run,false", "run,true", "doctor,false", "doctor,true"})
  void invalidAuthorityIsARejectedCommandWithOriginalFileAndCursor(String command, boolean file)
      throws java.io.IOException {
    Path request =
        Files.writeString(
            root.resolve("request.json"),
            """
            {"protocolVersion":1,"source":{"type":"NEW"},"persistence":{"type":"NONE"},
             "steps":[{"stepId":"ensure","target":{"type":"SHEET_BY_NAME","name":"Budget"},
             "action":{"type":"ENSURE_SHEET"}}]}
            """);
    Path grant = Files.writeString(root.resolve("authority.json"), "{\"maxInputSnapshotBytes\":}");
    Path response = root.resolve(command + "-response.json");
    var args =
        new ArrayList<>(
            List.of(command, "--request", request.toString(), "--grant", grant.toString()));
    if (file) args.addAll(List.of("--response", response.toString()));
    var stdout = new ByteArrayOutputStream();
    var stderr = new ByteArrayOutputStream();
    int exit =
        new GridGrindCli()
            .run(args.toArray(String[]::new), InputStream.nullInputStream(), stdout, stderr);
    assertEquals(2, exit);
    assertEquals(0, stderr.size());
    if (file) assertEquals(0, stdout.size());
    var error =
        GridGrindCliJson.readBytes(
            file ? Files.readAllBytes(response) : stdout.toByteArray(), CommandError.class);
    assertEquals(GridGrindProblemCode.INVALID_AUTHORITY, error.problems().getFirst().code());
    var context =
        assertInstanceOf(ProblemContext.ReadAuthority.class, error.problems().getFirst().context());
    assertEquals(grant.toString(), context.authorityPath());
    assertTrue(context.json().byteOffsetValue().isPresent());
  }
}
