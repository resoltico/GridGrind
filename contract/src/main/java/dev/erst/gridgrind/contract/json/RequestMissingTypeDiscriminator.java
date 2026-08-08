package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reports a missing discriminator required to choose a sealed input variant. */
record RequestMissingTypeDiscriminator(String path) implements RequestShapeStructuralProblem {
  RequestMissingTypeDiscriminator {
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
