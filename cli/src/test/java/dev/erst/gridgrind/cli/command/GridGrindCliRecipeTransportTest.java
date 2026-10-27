package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Transport-boundary regressions for printable recipe payloads. */
class GridGrindCliRecipeTransportTest extends GridGrindCliTestSupport {
  @Test
  void successfulDiscoveryCommandsReportResponseTransportFallbacksAsFailures() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-recipe-response-directory-");

    assertResponseTransportFailure(
        new String[] {"recipe", "show", "BUDGET", "--response", responseDirectory.toString()});
    assertResponseTransportFailure(
        new String[] {"recipe", "list", "--response", responseDirectory.toString()});
    assertResponseTransportFailure(
        new String[] {
          "recipe", "show", "BUDGET", "--details", "--response", responseDirectory.toString()
        });
    assertResponseTransportFailure(
        new String[] {"recipe", "match", "budget", "--response", responseDirectory.toString()});
  }

  @Test
  void assetBackedRecipeResponseCollisionRejectsBeforeMaterializingAssets() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-recipe-response-collision-", ".json");
    Files.writeString(responsePath, "sentinel\n");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    GridGrindCli cli = new GridGrindCli();

    try {
      int exitCode =
          cli.run(
              new String[] {
                "recipe",
                "show",
                "PACKAGE_SECURITY_INSPECTION",
                "--response",
                responsePath.toString()
              },
              InputStream.nullInputStream(),
              stdout,
              stderr);
      assertEquals(2, exitCode);
      assertEquals(
          dev.erst.gridgrind.core.problem.GridGrindProblemCode.INVALID_ARGUMENTS,
          commandErrorOnStdout(stdout, stderr).primaryProblem().code());
      assertEquals("sentinel\n", Files.readString(responsePath));
    } finally {
      Files.deleteIfExists(responsePath);
    }
  }

  private static void assertResponseTransportFailure(String[] arguments) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode = new GridGrindCli().run(arguments, InputStream.nullInputStream(), stdout, stderr);

    assertEquals(1, exitCode);
    assertTrue(stdout.size() > 0, "fallback must preserve the discovery payload on stdout");
    assertTrue(stderr.size() > 0, "fallback must report the unavailable response destination");
  }
}
