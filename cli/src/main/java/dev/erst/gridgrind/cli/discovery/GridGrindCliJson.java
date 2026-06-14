package dev.erst.gridgrind.cli.discovery;

import java.io.IOException;
import java.io.OutputStream;

/** Shared JSON codec for CLI-owned discovery surfaces. */
public final class GridGrindCliJson {
  private GridGrindCliJson() {}

  /** Reads one typed CLI discovery payload from UTF-8 JSON bytes. */
  public static <T> T readBytes(byte[] bytes, Class<T> valueType) throws IOException {
    return GridGrindCliJsonCodecSupport.readBytes(bytes, valueType);
  }

  /** Writes one CLI discovery payload to UTF-8 JSON bytes. */
  public static byte[] writeBytes(Object value) throws IOException {
    return GridGrindCliJsonCodecSupport.writeBytes(value);
  }

  /** Writes one CLI discovery payload as JSON without closing the caller-owned output stream. */
  public static void writeValue(OutputStream outputStream, Object value) throws IOException {
    GridGrindCliJsonCodecSupport.writeValue(outputStream, value);
  }
}
