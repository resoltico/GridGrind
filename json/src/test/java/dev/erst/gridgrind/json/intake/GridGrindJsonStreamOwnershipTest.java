package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Locks caller ownership of request streams across every request-analysis operation. */
class GridGrindJsonStreamOwnershipTest {
  @Test
  void requestInputStreamsRetainCallerOwnershipAcrossDecodeAndAnalysis() throws IOException {
    byte[] request =
        """
        {"protocolVersion":1,"source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[]}
        """
            .getBytes(StandardCharsets.UTF_8);
    try (TrackingInputStream decodeStream = new TrackingInputStream(request);
        TrackingInputStream analysisStream = new TrackingInputStream(request)) {
      assertEquals(
          ProtocolVersion.CURRENT, GridGrindJson.readRequest(decodeStream).protocolVersion());
      assertTrue(GridGrindJson.analyzeRequest(analysisStream).isStructurallyValid());
      assertFalse(decodeStream.closed);
      assertFalse(analysisStream.closed);
    }
  }

  /** Input stream that records whether the code under test closed its caller-owned stream. */
  private static final class TrackingInputStream extends ByteArrayInputStream {
    private boolean closed;

    private TrackingInputStream(byte[] payload) {
      super(payload);
    }

    @Override
    public void close() throws IOException {
      closed = true;
      super.close();
    }
  }
}
