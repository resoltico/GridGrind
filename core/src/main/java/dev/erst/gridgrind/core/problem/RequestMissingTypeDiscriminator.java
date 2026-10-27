package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Optional;

/** Reports a missing discriminator required to choose a sealed input variant. */
public record RequestMissingTypeDiscriminator(String path)
    implements RequestShapeStructuralProblem {
  public RequestMissingTypeDiscriminator {
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
