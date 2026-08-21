package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reports a value whose JSON token kind cannot satisfy its creator component. */
record RequestMalformedScalar(String path, String expected, long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  RequestMalformedScalar {
    path = RequestStructuralProblemSupport.requireText(path, "path");
    expected = RequestStructuralProblemSupport.requireText(expected, "expected");
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
    return "Field '" + path + "' must be " + expected;
  }
}
