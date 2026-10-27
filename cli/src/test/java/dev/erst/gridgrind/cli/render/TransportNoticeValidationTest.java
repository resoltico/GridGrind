package dev.erst.gridgrind.cli.render;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Transport rendering keeps its own optional response-path boundary. */
class TransportNoticeValidationTest {
  @Test
  void rejectsBlankResponsePath() {
    var failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new CliTransportNotice(
                    CliTransportNotice.Reason.RESPONSE_WRITE_FAILED,
                    CliTransportNotice.Destination.STDOUT,
                    Optional.of("   ")));
    assertEquals("responsePath must not be blank", failure.getMessage());
  }
}
