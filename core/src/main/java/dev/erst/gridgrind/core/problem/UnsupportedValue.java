package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Optional;

/** One scalar value was outside the allowed vocabulary for a field. */
public record UnsupportedValue(
    String value, Optional<String> jsonPathValue, List<String> allowedValues)
    implements RequestProblemDescriptor.Shape {
  public UnsupportedValue {
    value = TextValidation.requireNonBlank(value, "value");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
    allowedValues = RequestProblemDescriptorSupport.copyStrings(allowedValues, "allowedValues");
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
