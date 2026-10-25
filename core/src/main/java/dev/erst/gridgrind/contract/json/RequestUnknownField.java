package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;

/** Reports an object member that is not part of the applicable creator contract. */
record RequestUnknownField(String path, long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  RequestUnknownField {
    Objects.requireNonNull(path, "path must not be null");
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
  }

  @Override
  public Optional<String> jsonPath() {
    return RequestStructuralProblemSupport.optionalJsonPath(path);
  }

  @Override
  public Optional<Long> byteOffset() {
    return Optional.of(tokenByteOffset);
  }

  @Override
  public String message() {
    return "Unknown field '" + path + "'";
  }
}
