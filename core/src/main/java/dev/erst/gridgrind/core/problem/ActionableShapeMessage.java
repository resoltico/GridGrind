package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Optional;

/** Public request-shape message whose throw site also owns one specific repair sentence. */
public record ActionableShapeMessage(
    String message, String resolutionValue, Optional<String> jsonPathValue)
    implements RequestProblemDescriptor.Shape {
  public ActionableShapeMessage {
    message = TextValidation.requireNonBlank(message, "message");
    resolutionValue = TextValidation.requireNonBlank(resolutionValue, "resolutionValue");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
