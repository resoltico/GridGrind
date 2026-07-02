package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** One authored field was explicitly null instead of omitted. */
public record ExplicitNullField(String jsonPathValue) implements RequestProblemDescriptor.Shape {
  public ExplicitNullField {
    jsonPathValue = RequestProblemDescriptorSupport.requireJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(jsonPathValue);
  }
}
