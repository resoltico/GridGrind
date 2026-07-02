package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.ByteArrayOutputStream;
import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;
import tools.jackson.databind.DeserializationFeature;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Shared JSON mapper and stream utilities for CLI-owned discovery payloads. */
final class GridGrindCliJsonCodecSupport {
  static final JsonMapper JSON_MAPPER =
      JsonMapper.builder().enable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES).build();

  private GridGrindCliJsonCodecSupport() {}

  static <T> T readBytes(byte[] bytes, Class<T> valueType) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    Objects.requireNonNull(valueType, "valueType must not be null");
    return JSON_MAPPER.readValue(bytes, valueType);
  }

  static <T> T readStream(InputStream inputStream, Class<T> valueType) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    Objects.requireNonNull(valueType, "valueType must not be null");
    return JSON_MAPPER.readValue(nonClosing(inputStream), valueType);
  }

  static JsonNode readTree(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return JSON_MAPPER.readTree(bytes);
  }

  static byte[] writeBytes(Object value) throws IOException {
    return writeBytes(value, false);
  }

  static byte[] writeBytes(Object value, boolean pretty) throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    writeValue(outputStream, value, pretty);
    return outputStream.toByteArray();
  }

  static void writeValue(OutputStream outputStream, Object value) throws IOException {
    writeValue(outputStream, value, false);
  }

  static void writeValue(OutputStream outputStream, Object value, boolean pretty)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(value, "value must not be null");
    GridGrindJsonOutput.writeCatalogLookupValue(outputStream, value, pretty);
  }

  private static InputStream nonClosing(InputStream inputStream) {
    return new FilterInputStream(inputStream) {
      @Override
      public void close() {
        // Caller owns the lifecycle of streams passed into the CLI discovery codec.
      }
    };
  }
}
