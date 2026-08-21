package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;

/** Reports one duplicate occurrence of a property name in an object. */
public record RequestDuplicateKey(
    String containingObjectPath, String key, int occurrenceOrdinal, long tokenByteOffset)
    implements RequestStructuralProblem {
  public RequestDuplicateKey {
    Objects.requireNonNull(containingObjectPath, "containingObjectPath must not be null");
    key = RequestStructuralProblemSupport.requireText(key, "key");
    if (occurrenceOrdinal < 0) {
      throw new IllegalArgumentException("occurrenceOrdinal must not be negative");
    }
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.empty();
  }

  @Override
  public Optional<Long> byteOffset() {
    return Optional.of(tokenByteOffset);
  }

  @Override
  public String message() {
    return "Duplicate key '" + key + "'";
  }
}
