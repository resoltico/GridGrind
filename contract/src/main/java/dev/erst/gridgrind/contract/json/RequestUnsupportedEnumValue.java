package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Optional;

/** Reports a scalar enum token outside the record's declared wire vocabulary. */
record RequestUnsupportedEnumValue(
    String path, String value, List<String> allowedValues, long tokenByteOffset)
    implements RequestShapeStructuralProblem {
  RequestUnsupportedEnumValue {
    path = RequestStructuralProblemSupport.requireText(path, "path");
    value = RequestStructuralProblemSupport.requireText(value, "value");
    allowedValues = RequestStructuralProblemSupport.copyStrings(allowedValues, "allowedValues");
    if (allowedValues.isEmpty()) {
      throw new IllegalArgumentException("allowedValues must not be empty");
    }
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
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
    return "Unsupported value '"
        + value
        + "' for field '"
        + path
        + "'; expected one of: "
        + String.join(", ", allowedValues);
  }
}
