package dev.erst.gridgrind.json.codec;

import java.util.Objects;

/** Unclassified authority reader defects remain unexpected, including argument exceptions. */
final class UnexpectedAuthorityReadException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  UnexpectedAuthorityReadException(RuntimeException cause) {
    super("Unexpected authority reader failure", Objects.requireNonNull(cause));
  }
}
