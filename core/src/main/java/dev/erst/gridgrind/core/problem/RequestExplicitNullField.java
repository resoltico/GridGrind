package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** Reports explicit null where omission is the only representation of absence. */
public record RequestExplicitNullField(String path, long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  public RequestExplicitNullField {
    path = RequestStructuralProblemSupport.requireText(path, "path");
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(path);
  }

  @Override
  public Optional<Long> byteOffset() {
    return Optional.of(tokenByteOffset);
  }

  @Override
  public String message() {
    return "Field '" + path + "' must be omitted when absent; explicit null is not accepted.";
  }
}
