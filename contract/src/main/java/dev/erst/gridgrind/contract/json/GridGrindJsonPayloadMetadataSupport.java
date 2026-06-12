package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Objects;
import java.util.Optional;
import tools.jackson.core.JacksonException;
import tools.jackson.core.TokenStreamLocation;

/** Owns JSON payload path and location extraction for contract error reporting. */
final class GridGrindJsonPayloadMetadataSupport {
  private GridGrindJsonPayloadMetadataSupport() {}

  static Optional<Integer> jsonLine(TokenStreamLocation location) {
    if (location == null) {
      return Optional.empty();
    }
    int line = location.getLineNr();
    return line > 0 ? Optional.of(line) : Optional.empty();
  }

  static Optional<Integer> jsonColumn(TokenStreamLocation location) {
    if (location == null) {
      return Optional.empty();
    }
    int column = location.getColumnNr();
    return column > 0 ? Optional.of(column) : Optional.empty();
  }

  static PayloadMetadata payloadMetadata(JacksonException exception) {
    return new PayloadMetadata(
        jsonPath(exception),
        jsonLine(exception.getLocation()),
        jsonColumn(exception.getLocation()));
  }

  static Optional<String> jsonPath(JacksonException exception) {
    String rendered = renderPath(exception.getPath());
    return rendered.isEmpty() ? Optional.empty() : Optional.of(rendered);
  }

  static Optional<String> terminalContainerName(List<JacksonException.Reference> path) {
    if (path.isEmpty()) {
      return Optional.empty();
    }
    return Optional.ofNullable(path.getLast().getPropertyName());
  }

  static String renderPath(List<JacksonException.Reference> path) {
    StringBuilder rendered = new StringBuilder();
    for (JacksonException.Reference reference : path) {
      String name = reference.getPropertyName();
      if (name != null) {
        if (!rendered.isEmpty()) {
          rendered.append('.');
        }
        rendered.append(name);
      } else {
        rendered.append('[').append(reference.getIndex()).append(']');
      }
    }
    return rendered.toString();
  }

  record PayloadMetadata(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    PayloadMetadata {
      Objects.requireNonNull(jsonPath, "jsonPath must not be null");
      Objects.requireNonNull(jsonLine, "jsonLine must not be null");
      Objects.requireNonNull(jsonColumn, "jsonColumn must not be null");
    }
  }
}
