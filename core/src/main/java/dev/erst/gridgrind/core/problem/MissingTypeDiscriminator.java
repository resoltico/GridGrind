package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** One required discriminator field was absent. */
public record MissingTypeDiscriminator(String jsonPathValue)
    implements RequestProblemDescriptor.Shape {
  public MissingTypeDiscriminator {
    jsonPathValue = RequestProblemDescriptorSupport.requireJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(jsonPathValue);
  }
}
