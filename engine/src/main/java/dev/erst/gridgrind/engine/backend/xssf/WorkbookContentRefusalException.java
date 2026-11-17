package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.problem.WorkbookContentLocation;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** A known content refusal identified by the native reader or preservation operation. */
@SuppressWarnings("PMD.NonSerializableClass") // Native content diagnostics are process-local.
public final class WorkbookContentRefusalException extends RuntimeException {
  private static final long serialVersionUID = 1L;
  private final WorkbookContentIssue issue;
  private final Optional<WorkbookContentLocation> location;

  /** Retains the native cause privately while exposing only the core's safe refusal wording. */
  public WorkbookContentRefusalException(
      WorkbookContentIssue issue,
      Optional<WorkbookContentLocation> location,
      @Nullable Throwable cause) {
    super(Objects.requireNonNull(issue, "issue must not be null").message(), cause);
    this.issue = issue;
    this.location = Objects.requireNonNull(location, "location must not be null");
  }

  /** Returns the core-owned refusal reason and kind. */
  public WorkbookContentIssue issue() {
    return issue;
  }

  /** Returns only source coordinates established by the native operation. */
  public Optional<WorkbookContentLocation> location() {
    return location;
  }
}
