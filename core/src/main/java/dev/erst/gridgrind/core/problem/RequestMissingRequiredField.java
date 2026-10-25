package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** Reports one omitted creator field that has no protocol default. */
public record RequestMissingRequiredField(String path) implements RequestShapeStructuralProblem {
  public RequestMissingRequiredField {
    path = RequestStructuralProblemSupport.requireText(path, "path");
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
