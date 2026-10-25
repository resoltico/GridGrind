package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** One authored step id duplicated an earlier step id. */
public record DuplicateStepId(dev.erst.gridgrind.core.vocabulary.StepId value, String jsonPathValue)
    implements RequestProblemDescriptor.Invariant {
  public DuplicateStepId {
    java.util.Objects.requireNonNull(value, "value must not be null");
    jsonPathValue = RequestProblemDescriptorSupport.requireJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(jsonPathValue);
  }
}
