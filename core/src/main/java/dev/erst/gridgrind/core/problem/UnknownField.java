package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** One unexpected request field was authored. */
public record UnknownField(String jsonPathValue) implements RequestProblemDescriptor.Shape {
  public UnknownField {
    jsonPathValue = RequestProblemDescriptorSupport.requireJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(jsonPathValue);
  }
}
