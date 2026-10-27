package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Optional;

/** Public request-invariant message whose throw site also owns one specific repair sentence. */
public record ActionableInvariantMessage(
    String message, String resolutionValue, Optional<String> jsonPathValue)
    implements RequestProblemDescriptor.Invariant {
  public ActionableInvariantMessage {
    message = TextValidation.requireNonBlank(message, "message");
    resolutionValue = TextValidation.requireNonBlank(resolutionValue, "resolutionValue");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
