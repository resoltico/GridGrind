package dev.erst.gridgrind.contract.json;

import java.io.Serializable;
import java.util.Objects;
import java.util.Optional;

/** Serializable JSON-location value used by payload exceptions without nullable field padding. */
public sealed interface PayloadLocation extends Serializable
    permits PayloadLocation.Unavailable,
        PayloadLocation.PathOnly,
        PayloadLocation.LineColumn,
        PayloadLocation.Located {
  /** Collapses explicit optional path and cursor inputs into one serializable location variant. */
  static PayloadLocation from(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    Optional<String> normalizedPath = Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    Optional<Integer> normalizedLine =
        Objects.requireNonNull(jsonLine, "jsonLine must not be null");
    Optional<Integer> normalizedColumn =
        Objects.requireNonNull(jsonColumn, "jsonColumn must not be null");
    if (normalizedLine.isPresent() != normalizedColumn.isPresent()) {
      return normalizedPath.isPresent()
          ? new PathOnly(normalizedPath.orElseThrow())
          : new Unavailable();
    }
    if (normalizedPath.isPresent() && normalizedLine.isPresent()) {
      return new Located(
          normalizedPath.orElseThrow(),
          normalizedLine.orElseThrow(),
          normalizedColumn.orElseThrow());
    }
    if (normalizedPath.isPresent()) {
      return new PathOnly(normalizedPath.orElseThrow());
    }
    if (normalizedLine.isPresent()) {
      return new LineColumn(normalizedLine.orElseThrow(), normalizedColumn.orElseThrow());
    }
    return new Unavailable();
  }

  /** Returns the JSON path when one request field location was identified. */
  Optional<String> jsonPath();

  /** Returns the request JSON line when the parser exposed one cursor. */
  Optional<Integer> jsonLine();

  /** Returns the request JSON column when the parser exposed one cursor. */
  Optional<Integer> jsonColumn();

  record Unavailable() implements PayloadLocation {
    @Override
    public Optional<String> jsonPath() {
      return Optional.empty();
    }

    @Override
    public Optional<Integer> jsonLine() {
      return Optional.empty();
    }

    @Override
    public Optional<Integer> jsonColumn() {
      return Optional.empty();
    }
  }

  record PathOnly(String jsonPathValue) implements PayloadLocation {
    public PathOnly {
      Objects.requireNonNull(jsonPathValue, "jsonPathValue must not be null");
      if (jsonPathValue.isBlank()) {
        throw new IllegalArgumentException("jsonPathValue must not be blank");
      }
    }

    @Override
    public Optional<String> jsonPath() {
      return Optional.of(jsonPathValue);
    }

    @Override
    public Optional<Integer> jsonLine() {
      return Optional.empty();
    }

    @Override
    public Optional<Integer> jsonColumn() {
      return Optional.empty();
    }
  }

  record LineColumn(int jsonLineValue, int jsonColumnValue) implements PayloadLocation {
    public LineColumn {
      if (jsonLineValue < 1) {
        throw new IllegalArgumentException("jsonLineValue must be greater than 0");
      }
      if (jsonColumnValue < 1) {
        throw new IllegalArgumentException("jsonColumnValue must be greater than 0");
      }
    }

    @Override
    public Optional<String> jsonPath() {
      return Optional.empty();
    }

    @Override
    public Optional<Integer> jsonLine() {
      return Optional.of(jsonLineValue);
    }

    @Override
    public Optional<Integer> jsonColumn() {
      return Optional.of(jsonColumnValue);
    }
  }

  record Located(String jsonPathValue, int jsonLineValue, int jsonColumnValue)
      implements PayloadLocation {
    public Located {
      Objects.requireNonNull(jsonPathValue, "jsonPathValue must not be null");
      if (jsonPathValue.isBlank()) {
        throw new IllegalArgumentException("jsonPathValue must not be blank");
      }
      if (jsonLineValue < 1) {
        throw new IllegalArgumentException("jsonLineValue must be greater than 0");
      }
      if (jsonColumnValue < 1) {
        throw new IllegalArgumentException("jsonColumnValue must be greater than 0");
      }
    }

    @Override
    public Optional<String> jsonPath() {
      return Optional.of(jsonPathValue);
    }

    @Override
    public Optional<Integer> jsonLine() {
      return Optional.of(jsonLineValue);
    }

    @Override
    public Optional<Integer> jsonColumn() {
      return Optional.of(jsonColumnValue);
    }
  }
}
