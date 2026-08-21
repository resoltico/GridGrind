package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reports request input that is not valid UTF-8. */
record RequestInvalidEncoding(String message, long tokenByteOffset)
    implements RequestStructuralProblem {
  RequestInvalidEncoding {
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
