package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
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
        {"protocolVersion":"V2","source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[]}
        """
            .getBytes(StandardCharsets.UTF_8);
    try (TrackingInputStream decodeStream = new TrackingInputStream(request);
        TrackingInputStream analysisStream = new TrackingInputStream(request)) {
      assertEquals(
          GridGrindProtocolVersion.V2, GridGrindJson.readRequest(decodeStream).protocolVersion());
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
