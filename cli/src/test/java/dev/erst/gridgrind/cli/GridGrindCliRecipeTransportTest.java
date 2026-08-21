package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Transport-boundary regressions for printable recipe payloads. */
class GridGrindCliRecipeTransportTest extends GridGrindCliTestSupport {
  @Test
  void assetBackedRecipeResponseCollisionPreservesTheOriginalRecipePayload() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-recipe-response-collision-", ".json");
    Files.writeString(responsePath, "sentinel\n");
    ByteArrayOutputStream expected = new ByteArrayOutputStream();
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    GridGrindCli cli = new GridGrindCli();

    try {
      int directExitCode =
          cli.run(
              new String[] {"--print-recipe", "--lookup", "PACKAGE_SECURITY_INSPECTION"},
              InputStream.nullInputStream(),
              expected);
      int exitCode =
          cli.run(
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
      CliTransportNotice notice =
          GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);

      assertEquals(0, directExitCode);
      assertEquals(1, exitCode);
      assertArrayEquals(expected.toByteArray(), stdout.toByteArray());
      assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
      assertEquals(
          java.util.Optional.of(responsePath.toAbsolutePath().toString()), notice.responsePath());
      assertEquals("sentinel\n", Files.readString(responsePath));
    } finally {
      Files.deleteIfExists(responsePath);
    }
  }
}
