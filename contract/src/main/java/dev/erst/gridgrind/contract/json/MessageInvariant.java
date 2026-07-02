package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Public request-invariant message whose semantics are already fully decided at the throw site. */
public record MessageInvariant(String message, Optional<String> jsonPathValue)
    implements RequestProblemDescriptor.Invariant {
  public MessageInvariant {
    message = RequestProblemDescriptorSupport.requireNonBlank(message, "message");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
