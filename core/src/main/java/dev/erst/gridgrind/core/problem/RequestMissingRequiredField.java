package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Optional;

/** Reports one omitted creator field that has no protocol default. */
public record RequestMissingRequiredField(String path) implements RequestShapeStructuralProblem {
  public RequestMissingRequiredField {
    path = TextValidation.requireNonBlank(path, "path");
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(path);
  }

  @Override
  public Optional<Long> byteOffset() {
    return Optional.empty();
  }

  @Override
  public String message() {
    return "Missing required field '" + path + "'";
  }
}
