package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** One authored step id duplicated an earlier step id. */
public record DuplicateStepId(String value, String jsonPathValue)
    implements RequestProblemDescriptor.Invariant {
  public DuplicateStepId {
    value = RequestProblemDescriptorSupport.requireNonBlank(value, "value");
    jsonPathValue = RequestProblemDescriptorSupport.requireJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(jsonPathValue);
  }
}
