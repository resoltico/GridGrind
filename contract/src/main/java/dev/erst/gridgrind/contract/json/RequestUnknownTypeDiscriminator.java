package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Optional;

/** Reports an unrecognised sealed-input discriminator. */
record RequestUnknownTypeDiscriminator(
    String path,
    String value,
    List<String> similarValues,
    Optional<String> specificGuidance,
    long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  RequestUnknownTypeDiscriminator {
    path = RequestStructuralProblemSupport.requireText(path, "path");
    value = RequestStructuralProblemSupport.requireText(value, "value");
    similarValues = RequestStructuralProblemSupport.copyStrings(similarValues, "similarValues");
    specificGuidance =
        RequestStructuralProblemSupport.copyOptionalText(specificGuidance, "specificGuidance");
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
  }

  RequestUnknownTypeDiscriminator(String path, String value, long tokenByteOffset) {
    this(path, value, List.of(), Optional.empty(), tokenByteOffset);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(path);
  }

  @Override
  public Optional<Long> byteOffset() {
    return Optional.of(tokenByteOffset);
  }

  @Override
  public String message() {
    StringBuilder message =
        new StringBuilder(64).append("Unknown type value '").append(value).append('\'');
    specificGuidance.ifPresent(guidance -> message.append("; ").append(guidance));
    if (!similarValues.isEmpty()) {
      message.append("; similar valid values: ").append(String.join(", ", similarValues));
    }
    return message.toString();
  }
}
