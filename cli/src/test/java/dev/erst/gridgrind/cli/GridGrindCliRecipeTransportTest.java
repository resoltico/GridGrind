package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Transport-boundary regressions for printable recipe payloads. */
class GridGrindCliRecipeTransportTest extends GridGrindCliTestSupport {
  @Test
  void assetBackedRecipeResponseCollisionEmitsOnlyTheStructuredTransportNotice()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-recipe-response-collision-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    try {
      int exitCode =
          new GridGrindCli()
              .run(
                  new String[] {
                    "--print-recipe",
                    "--lookup",
                    "PACKAGE_SECURITY_INSPECTION",
                    "--response",
                    responsePath.toString()
                  },
                  InputStream.nullInputStream(),
                  stdout,
                  stderr);

      CommandError failure = commandError(stdout.toByteArray());
      CliTransportNotice notice =
          GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);

      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.IO_ERROR, failure.primaryProblem().code());
      assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
      assertEquals(
          java.util.Optional.of(responsePath.toAbsolutePath().toString()), notice.responsePath());
    } finally {
      Files.deleteIfExists(responsePath);
    }
  }
}
