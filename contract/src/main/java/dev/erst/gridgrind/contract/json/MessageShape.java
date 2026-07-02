package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Public request-shape message whose semantics are already fully decided at the throw site. */
public record MessageShape(String message, Optional<String> jsonPathValue)
    implements RequestProblemDescriptor.Shape {
  public MessageShape {
    message = RequestProblemDescriptorSupport.requireNonBlank(message, "message");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
