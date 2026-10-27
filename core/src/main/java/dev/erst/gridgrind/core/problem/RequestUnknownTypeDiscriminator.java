package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Reports an unrecognised sealed-input discriminator. */
public record RequestUnknownTypeDiscriminator(
    String path,
    String value,
    List<String> similarValues,
    Optional<String> specificGuidance,
    long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  public RequestUnknownTypeDiscriminator {
    path = TextValidation.requireNonBlank(path, "path");
    Objects.requireNonNull(value, "value must not be null");
    similarValues = RequestStructuralProblemSupport.copyStrings(similarValues, "similarValues");
    specificGuidance =
        RequestStructuralProblemSupport.copyOptionalText(specificGuidance, "specificGuidance");
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
  }

  /** Creates an unknown-type finding without suggestions or extra guidance. */
  public RequestUnknownTypeDiscriminator(String path, String value, long tokenByteOffset) {
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
