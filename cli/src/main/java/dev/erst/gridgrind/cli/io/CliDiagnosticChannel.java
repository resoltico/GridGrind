package dev.erst.gridgrind.cli.io;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.Objects;

/** Invocation-owned serialization of complete stderr frames; the destination remains borrowed. */
public final class CliDiagnosticChannel extends OutputStream {
  private final OutputStream destination;
  private final java.util.concurrent.locks.ReentrantLock frameLock =
      new java.util.concurrent.locks.ReentrantLock();

  /** Wraps the invocation's stderr without changing process streams or acquiring a console lock. */
  public CliDiagnosticChannel(OutputStream destination) {
    this.destination = Objects.requireNonNull(destination);
  }

  /** Writes exactly one newline-terminated frame under this channel's private lock. */
  public void writeFrame(byte[] payload) throws IOException {
    frameLock.lock();
    try {
      destination.write(payload);
      if (payload.length == 0 || payload[payload.length - 1] != '\n') destination.write('\n');
      flush();
    } finally {
      frameLock.unlock();
    }
  }

  @Override
  public void write(int value) throws IOException {
    frameLock.lock();
    try {
      destination.write(value);
    } finally {
      frameLock.unlock();
    }
  }

  @Override
  public void write(byte[] bytes, int offset, int length) throws IOException {
    frameLock.lock();
    try {
      destination.write(bytes, offset, length);
    } finally {
      frameLock.unlock();
    }
  }

  @Override
  @SuppressWarnings("PMD.CloseResource") // The destination PrintStream is borrowed.
  public void flush() throws IOException {
    frameLock.lock();
    try {
      destination.flush();
      if (destination instanceof PrintStream console && console.checkError()) {
        throw new IOException("CLI diagnostic stream did not accept the frame.");
      }
    } finally {
      frameLock.unlock();
    }
  }

  @Override
  public void close() throws IOException {
    flush();
  }
}
