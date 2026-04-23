package dev.erst.gridgrind.cli;

import java.io.ByteArrayInputStream;
import java.io.IOException;

/** Shared helpers for CLI integration tests. */
class GridGrindCliTestSupport {
  protected GridGrindCliTestSupport() {}

  /** ByteArrayInputStream that records whether {@code close()} was called. */
  protected static final class TrackingInputStream extends ByteArrayInputStream {
    private boolean closed;

    TrackingInputStream(byte[] bytes) {
      super(bytes);
    }

    @Override
    public void close() throws IOException {
      closed = true;
      super.close();
    }

    boolean closed() {
      return closed;
    }
  }
}
