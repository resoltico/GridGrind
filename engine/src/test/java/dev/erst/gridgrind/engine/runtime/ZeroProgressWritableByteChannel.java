package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.WritableByteChannel;
import java.util.Objects;

/** Reports zero bytes for its first write to force the publication transfer fallback. */
final class ZeroProgressWritableByteChannel implements WritableByteChannel {
  private final WritableByteChannel delegate;
  private boolean firstWrite = true;

  ZeroProgressWritableByteChannel(WritableByteChannel delegate) {
    this.delegate = Objects.requireNonNull(delegate, "delegate must not be null");
  }

  @Override
  public int write(ByteBuffer source) throws IOException {
    if (firstWrite) {
      firstWrite = false;
      return 0;
    }
    return delegate.write(source);
  }

  @Override
  public boolean isOpen() {
    return delegate.isOpen();
  }

  @Override
  public void close() {
    // The enclosing test owns the delegate channel.
  }
}
