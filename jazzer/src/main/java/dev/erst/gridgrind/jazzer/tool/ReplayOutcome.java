package dev.erst.gridgrind.jazzer.tool;

import java.util.Objects;
import java.util.Optional;

/** Captures whether replaying a local Jazzer input succeeded, was expected-invalid, or failed. */
public sealed interface ReplayOutcome
    permits ReplayOutcome.ExpectedInvalid, ReplayOutcome.Success, ReplayOutcome.UnexpectedFailure {
  /** Returns the harness key that produced this replay outcome. */
  String harnessKey();

  /** Returns the structured details decoded from the raw input bytes. */
  ReplayDetails details();

  /** Represents a replay that completed without surfacing a bug. */
  record Success(String harnessKey, ReplayDetails details) implements ReplayOutcome {}

  /** Represents a replay that was invalid in a documented, expected way. */
  record ExpectedInvalid(
      String harnessKey, String invalidKind, Optional<String> message, ReplayDetails details)
      implements ReplayOutcome {
    public ExpectedInvalid {
      Objects.requireNonNull(harnessKey, "harnessKey must not be null");
      Objects.requireNonNull(invalidKind, "invalidKind must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(details, "details must not be null");
    }
  }

  /** Represents a replay that surfaced an unexpected exception or invariant failure. */
  record UnexpectedFailure(
      String harnessKey,
      String failureKind,
      Optional<String> message,
      String stackTrace,
      ReplayDetails details)
      implements ReplayOutcome {
    public UnexpectedFailure {
      Objects.requireNonNull(harnessKey, "harnessKey must not be null");
      Objects.requireNonNull(failureKind, "failureKind must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(stackTrace, "stackTrace must not be null");
      Objects.requireNonNull(details, "details must not be null");
    }
  }
}
