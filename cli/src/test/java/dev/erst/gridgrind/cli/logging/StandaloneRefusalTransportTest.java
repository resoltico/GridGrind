package dev.erst.gridgrind.cli.logging;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.cli.command.CliStandaloneFailureSupport;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;
import org.junit.jupiter.api.Test;

/** Failed primary transport during preinitialization remains a quiet nonzero process outcome. */
class StandaloneRefusalTransportTest {
  @Test
  void rejectedLoggingPolicyDoesNotEscapeThroughABrokenPrimaryOutput() throws IOException {
    var refusal =
        assertThrows(
            StandaloneLoggingRefusal.class,
            () ->
                StandaloneLoggingPreflight.level(
                    Map.of("log4j2.configurationFile", "private-value")));
    var stderr = new ByteArrayOutputStream();
    try (var stdout =
        new OutputStream() {
          @Override
          public void write(int value) throws IOException {
            throw new IOException("closed pipe");
          }
        }) {
      assertEquals(
          1, CliStandaloneFailureSupport.emit(new String[] {"version"}, stdout, stderr, refusal));
      assertEquals(0, stderr.size());
    }
  }
}
