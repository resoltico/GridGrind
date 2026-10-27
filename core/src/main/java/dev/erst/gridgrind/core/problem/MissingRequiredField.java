package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** One required request field was absent. */
public record MissingRequiredField(String jsonPathValue) implements RequestProblemDescriptor.Shape {
  public MissingRequiredField {
    jsonPathValue = RequestProblemDescriptorSupport.requireJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(jsonPathValue);
  }
}
