package dev.erst.gridgrind.jazzer.tool;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.SerializationFeature;

/** Centralizes JSON serialization for the local Jazzer operator layer. */
public final class JazzerJson {
  private static final JsonMapper JSON_MAPPER =
      JsonMapper.builder().enable(SerializationFeature.INDENT_OUTPUT).build();

  private JazzerJson() {}

  /** Writes one value as pretty-printed JSON to the requested path. */
  public static void write(Path path, Object value) throws IOException {
    Objects.requireNonNull(path, "path must not be null");
    Objects.requireNonNull(value, "value must not be null");
    JSON_MAPPER.writeValue(path.toFile(), value);
  }

  /** Reads one JSON value from disk into the requested type. */
  public static <T> T read(Path path, Class<T> type) throws IOException {
    Objects.requireNonNull(path, "path must not be null");
    Objects.requireNonNull(type, "type must not be null");
    return JSON_MAPPER.readValue(path.toFile(), type);
  }

  /** Reads one JSON value from a classpath resource into the requested type. */
  public static <T> T readResource(String resourcePath, Class<T> type) throws IOException {
    Objects.requireNonNull(resourcePath, "resourcePath must not be null");
    Objects.requireNonNull(type, "type must not be null");
    try (InputStream inputStream = JazzerJson.class.getResourceAsStream(resourcePath)) {
      if (inputStream == null) {
        throw new IOException("Classpath resource does not exist: " + resourcePath);
      }
      return JSON_MAPPER.readValue(inputStream, type);
    }
  }

  /** Returns one value as pretty-printed JSON text. */
  public static String toJson(Object value) throws IOException {
    Objects.requireNonNull(value, "value must not be null");
    return JSON_MAPPER.writeValueAsString(value);
  }
}
