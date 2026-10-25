package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** Reports request input that is not valid UTF-8. */
public record RequestInvalidEncoding(String message, long tokenByteOffset)
    implements RequestStructuralProblem {
  public RequestInvalidEncoding {
    message = RequestStructuralProblemSupport.requireText(message, "message");
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.empty();
  }

  @Override
  public Optional<Long> byteOffset() {
    return Optional.of(tokenByteOffset);
  }
}
