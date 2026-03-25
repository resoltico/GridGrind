package dev.erst.gridgrind.protocol;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.DateTimeException;
import java.util.Objects;
import tools.jackson.core.JacksonException;
import tools.jackson.core.StreamReadFeature;
import tools.jackson.core.StreamWriteFeature;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.databind.SerializationFeature;
import tools.jackson.databind.json.JsonMapper;

/** Shared JSON codec for the GridGrind protocol. */
public final class GridGrindJson {
  private static final JsonMapper JSON_MAPPER =
      JsonMapper.builder()
          .disable(StreamReadFeature.AUTO_CLOSE_SOURCE)
          .disable(StreamWriteFeature.AUTO_CLOSE_TARGET)
          .enable(SerializationFeature.INDENT_OUTPUT)
          .build();

  private GridGrindJson() {}

  /** Reads a request from an input stream without closing the caller-owned stream. */
  public static GridGrindRequest readRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, GridGrindRequest.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a request from a byte array. */
  public static GridGrindRequest readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, GridGrindRequest.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static GridGrindResponse readResponse(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, GridGrindResponse.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a response from a byte array. */
  public static GridGrindResponse readResponse(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, GridGrindResponse.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Serializes a response to bytes. */
  public static byte[] writeResponseBytes(GridGrindResponse response) throws IOException {
    Objects.requireNonNull(response, "response must not be null");
    return JSON_MAPPER.writeValueAsBytes(response);
  }

  /** Writes a response to an output stream without closing the caller-owned stream. */
  public static void writeResponse(OutputStream outputStream, GridGrindResponse response)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    outputStream.write(writeResponseBytes(response));
  }

  private static IllegalArgumentException invalidPayload(JacksonException exception) {
    PayloadMetadata metadata = payloadMetadata(exception);
    Throwable validationCause = validationCause(exception);
    if (validationCause != null) {
      return new InvalidRequestException(
          message(validationCause),
          metadata.jsonPath(),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    return new InvalidJsonException(
        message(exception),
        metadata.jsonPath(),
        metadata.jsonLine(),
        metadata.jsonColumn(),
        exception);
  }

  private static Throwable validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof IllegalArgumentException || current instanceof DateTimeException) {
        return current;
      }
      current = current.getCause();
    }
    return null;
  }

  private static String message(Throwable throwable) {
    String message =
        throwable instanceof JacksonException jacksonException
            ? jacksonException.getOriginalMessage()
            : throwable.getMessage();
    return cleanJacksonMessage(message);
  }

  private static String cleanJacksonMessage(String message) {
    int startMarkerIndex = message.indexOf(" (start marker at [Source:");
    if (startMarkerIndex >= 0) {
      return message.substring(0, startMarkerIndex);
    }
    return message;
  }

  private static PayloadMetadata payloadMetadata(JacksonException exception) {
    return new PayloadMetadata(
        jsonPath(exception),
        jsonLine(exception.getLocation()),
        jsonColumn(exception.getLocation()));
  }

  private static String jsonPath(JacksonException exception) {
    StringBuilder path = new StringBuilder();
    for (JacksonException.Reference reference : exception.getPath()) {
      String propertyName = reference.getPropertyName();
      if (propertyName != null) {
        if (!path.isEmpty()) {
          path.append('.');
        }
        path.append(propertyName);
        continue;
      }
      path.append('[').append(reference.getIndex()).append(']');
    }
    return path.isEmpty() ? null : path.toString();
  }

  /** Returns the 1-based line number from the location, or null when unavailable. */
  static Integer jsonLine(TokenStreamLocation location) {
    if (location == null) {
      return null;
    }
    int line = location.getLineNr();
    return line > 0 ? line : null;
  }

  /** Returns the 1-based column number from the location, or null when unavailable. */
  static Integer jsonColumn(TokenStreamLocation location) {
    if (location == null) {
      return null;
    }
    int column = location.getColumnNr();
    return column > 0 ? column : null;
  }

  private record PayloadMetadata(String jsonPath, Integer jsonLine, Integer jsonColumn) {}
}
