package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.io.CliDiagnosticChannel;
import dev.erst.gridgrind.cli.logging.StandaloneLogging;
import dev.erst.gridgrind.cli.logging.StandaloneLoggingPreflight;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.status.StatusLogger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Real provider/status routing is qualified in a child process without changing test-host policy.
 */
class StandaloneLoggingProcessTest {
  @TempDir Path root;

  @Test
  void nativeAndStatusEventsUseOnlyTheOwnedBoundedTransport() throws Exception {
    var result =
        GoldenLauncher.runProcess(
            List.of(
                Path.of(System.getProperty("java.home"), "bin", "java").toString(),
                "-cp",
                System.getProperty("java.class.path"),
                OwnedRouting.class.getName()),
            root,
            30);
    assertEquals(0, result.exit(), result.stdout() + result.stderr());
    assertEquals("", result.stderr());
    var lines = result.stdout().lines().toList();
    assertEquals(33, lines.size());
    assertTrue(lines.getFirst().contains("\"component\":\"LOGGING\""));
    assertTrue(lines.get(1).contains("\"level\":\"FATAL\""));
    assertTrue(lines.getLast().contains("NATIVE_DIAGNOSTIC_LIMIT"));
    assertFalse(result.stdout().contains("synthetic-private"));
  }

  /** Process-only fixture that deliberately emits status, custom-level and flood events. */
  @SuppressWarnings("PMD.PublicMemberInNonPublicType")
  // The child JVM needs a public fixture entrypoint; the JUnit owner stays package-private.
  public static final class OwnedRouting {
    private OwnedRouting() {}

    @SuppressWarnings({"PMD.UseProperClassLoader", "PMD.UnusedLocalVariable", "PMD.SystemPrintln"})
    // The child qualifies App-owned provider resources and prints only its captured result frames.
    public static void main(String[] args) throws java.io.IOException {
      var bytes = new ByteArrayOutputStream();
      var loader = OwnedRouting.class.getClassLoader();
      String threshold = StandaloneLoggingPreflight.validate(loader);
      try (var logging =
          StandaloneLogging.open(loader, threshold, new CliDiagnosticChannel(bytes))) {
        StatusLogger.getLogger().warn("synthetic-private-status");
        LogManager.getLogger("org.apache.poi.qualification")
            .log(
                Level.forName("synthetic-private-level", 100),
                "synthetic-private-native",
                new IllegalArgumentException("synthetic-private-exception"));
        for (int index = 0; index < 40; index++)
          LogManager.getLogger("org.apache.xmlbeans.qualification").warn("synthetic-private-flood");
      }
      System.out.print(bytes.toString(StandardCharsets.UTF_8));
    }
  }
}
