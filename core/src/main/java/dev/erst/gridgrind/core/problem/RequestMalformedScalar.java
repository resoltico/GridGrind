package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Optional;

/** Reports a value whose JSON token kind cannot satisfy its creator component. */
public record RequestMalformedScalar(String path, String expected, long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  public RequestMalformedScalar {
    path = TextValidation.requireNonBlank(path, "path");
    expected = TextValidation.requireNonBlank(expected, "expected");
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
