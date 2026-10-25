package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reports explicit null where omission is the only representation of absence. */
record RequestExplicitNullField(String path, long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  RequestExplicitNullField {
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
