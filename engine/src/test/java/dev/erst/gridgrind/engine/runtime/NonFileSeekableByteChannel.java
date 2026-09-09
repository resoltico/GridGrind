package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.SeekableByteChannel;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;

/** A non-file channel that makes publication capability and close-failure behavior observable. */
final class NonFileSeekableByteChannel implements SeekableByteChannel {
  private final AtomicBoolean closed;
  private final boolean failOnClose;

  NonFileSeekableByteChannel(AtomicBoolean closed, boolean failOnClose) {
    this.closed = Objects.requireNonNull(closed, "closed must not be null");
    this.failOnClose = failOnClose;
  }

  @Override
  public int read(ByteBuffer destination) {
    return -1;
  }

  @Override
  public int write(ByteBuffer source) {
    return source.remaining();
  }

  @Override
  public long position() {
    return 0;
  }

  @Override
  public SeekableByteChannel position(long newPosition) {
    return this;
  }

  @Override
  public long size() {
    return 0;
  }

  @Override
  public SeekableByteChannel truncate(long size) {
    return this;
  }

  @Override
  public boolean isOpen() {
    return !closed.get();
  }

  @Override
  public void close() throws IOException {
    closed.set(true);
    if (failOnClose) {
      throw new IOException("synthetic close failure");
    }
  }
}
