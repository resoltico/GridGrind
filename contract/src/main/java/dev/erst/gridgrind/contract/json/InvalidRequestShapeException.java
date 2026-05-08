package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Signals that valid JSON did not match the GridGrind request or response payload shape. */
public final class InvalidRequestShapeException extends IllegalArgumentException
    implements PayloadException {
  private static final long serialVersionUID = 1L;

  private final @Nullable String jsonPath;
  private final @Nullable Integer jsonLine;
  private final @Nullable Integer jsonColumn;

  /** Creates the exception with the given message, JSON location, and binding cause. */
  public InvalidRequestShapeException(
      String message,
      Optional<String> jsonPath,
      Optional<Integer> jsonLine,
      Optional<Integer> jsonColumn,
      @Nullable Throwable cause) {
    super(message, cause);
    this.jsonPath = Objects.requireNonNull(jsonPath, "jsonPath must not be null").orElse(null);
    this.jsonLine = Objects.requireNonNull(jsonLine, "jsonLine must not be null").orElse(null);
    this.jsonColumn =
        Objects.requireNonNull(jsonColumn, "jsonColumn must not be null").orElse(null);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.ofNullable(jsonPath);
  }

  @Override
  public Optional<Integer> jsonLine() {
    return Optional.ofNullable(jsonLine);
  }

  @Override
  public Optional<Integer> jsonColumn() {
    return Optional.ofNullable(jsonColumn);
  }
}
