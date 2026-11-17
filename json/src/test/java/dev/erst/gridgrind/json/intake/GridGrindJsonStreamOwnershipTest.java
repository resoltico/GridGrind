package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Locks caller ownership of request streams across every request-analysis operation. */
class GridGrindJsonStreamOwnershipTest {
  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  @SuppressWarnings("PMD.CloseResource")
  // This borrowed synthetic stream owns no resource; close() is the failing ownership witness.
  void oversizedStreamsStopAfterOneOverflowByteWithoutClosing(boolean analyze) {
    var stream =
        new InputStream() {
          private long count;

          @Override
          public int read() {
            recordRead(1);
            return ' ';
          }

          @Override
          public int read(byte[] bytes, int offset, int length) {
            java.util.Objects.checkFromIndexSize(offset, length, bytes.length);
            recordRead(length);
            java.util.Arrays.fill(bytes, offset, offset + length, (byte) ' ');
            return length;
          }

          private void recordRead(int length) {
            count += length;
            if (count > GridGrindJson.maxRequestDocumentBytes() + 1) {
              throw new AssertionError("Request reader consumed beyond overflow sentinel");
            }
          }

          @Override
          public void close() {
            throw new AssertionError("Borrowed stream was closed");
          }
        };
    assertThrows(
        dev.erst.gridgrind.core.problem.InvalidRequestException.class,
        () -> {
          if (analyze) GridGrindJson.analyzeRequest(stream);
          else GridGrindJson.readRequest(stream);
        });
  }

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
