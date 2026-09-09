package dev.erst.gridgrind.cli;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.SeekableByteChannel;
import java.util.Objects;

/** Reports end-of-file immediately while retaining the delegate's authoritative file size. */
final class EndOfFileChannel implements SeekableByteChannel {
  private final SeekableByteChannel delegate;

  EndOfFileChannel(SeekableByteChannel delegate) {
    this.delegate = Objects.requireNonNull(delegate, "delegate must not be null");
  }

  @Override
  public int read(ByteBuffer destination) {
    return -1;
  }

  @Override
  public int write(ByteBuffer source) throws IOException {
    return delegate.write(source);
  }

  @Override
  public long position() throws IOException {
    return delegate.position();
  }

  @Override
  public SeekableByteChannel position(long newPosition) throws IOException {
    delegate.position(newPosition);
    return this;
  }

  @Override
  public long size() throws IOException {
    return delegate.size();
  }

  @Override
  public SeekableByteChannel truncate(long size) throws IOException {
    delegate.truncate(size);
    return this;
  }

  @Override
  public boolean isOpen() {
    return delegate.isOpen();
  }

  @Override
  public void close() throws IOException {
    delegate.close();
  }
}
