package dev.erst.gridgrind.core.problem;

import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Whole-document authority refusal; no partial or default authority is accepted. */
public final class InvalidAuthorityException extends IllegalArgumentException
    implements PayloadException {
  private static final long serialVersionUID = 1L;
  private final AuthorityDocumentIssue issue;
  private final ProblemContextRequestSurfaces.JsonLocation location;

  /** Retains a known input issue and reliable original JSON coordinates. */
  public InvalidAuthorityException(
      AuthorityDocumentIssue issue,
      ProblemContextRequestSurfaces.JsonLocation location,
      @Nullable Throwable cause) {
    super(Objects.requireNonNull(issue, "issue must not be null").message(), cause);
    this.issue = issue;
    this.location = Objects.requireNonNull(location, "location must not be null");
  }

  /** Returns the input-role reason independently of transport status or exit codes. */
  public AuthorityDocumentIssue issue() {
    return issue;
  }

  /** Returns the full original document cursor, including a byte offset when available. */
  public ProblemContextRequestSurfaces.JsonLocation documentLocation() {
    return location;
  }

  @Override
  public PayloadLocation jsonLocation() {
    return PayloadLocation.from(
        location.jsonPathValue(), location.jsonLineValue(), location.jsonColumnValue());
  }
}
