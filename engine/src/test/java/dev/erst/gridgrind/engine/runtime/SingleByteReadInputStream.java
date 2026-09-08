package dev.erst.gridgrind.engine.runtime;

import java.io.InputStream;
import java.util.Arrays;

/** Delivers one byte per bulk read to exercise differing stream-progress behavior. */
final class SingleByteReadInputStream extends InputStream {
  private final byte[] bytes;
  private int index;

  SingleByteReadInputStream(byte[] bytes) {
    this.bytes = Arrays.copyOf(bytes, bytes.length);
  }

  @Override
  public int read(byte[] buffer, int offset, int length) {
    if (index >= bytes.length) {
      return -1;
    }
    if (length == 0) {
      return 0;
    }
    buffer[offset] = bytes[index];
    index++;
    return 1;
  }

  @Override
  public int read() {
    if (index >= bytes.length) {
      return -1;
    }
    int byteValue = bytes[index];
    index++;
    return byteValue;
  }
}
