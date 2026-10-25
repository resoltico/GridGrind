package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Public request-shape message whose throw site also owns one specific repair sentence. */
public record ActionableShapeMessage(
    String message, String resolutionValue, Optional<String> jsonPathValue)
    implements RequestProblemDescriptor.Shape {
  public ActionableShapeMessage {
    message = RequestProblemDescriptorSupport.requireNonBlank(message, "message");
    resolutionValue =
        RequestProblemDescriptorSupport.requireNonBlank(resolutionValue, "resolutionValue");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
