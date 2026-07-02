package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Optional;

/** One subtype discriminator value was unknown to the protocol. */
public record UnknownTypeValue(
    String typeId,
    Optional<String> jsonPathValue,
    List<String> similarValues,
    Optional<String> specificGuidance)
    implements RequestProblemDescriptor.Shape {
  public UnknownTypeValue {
    typeId = RequestProblemDescriptorSupport.requireNonBlank(typeId, "typeId");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
    similarValues = RequestProblemDescriptorSupport.copyStrings(similarValues, "similarValues");
    specificGuidance =
        RequestProblemDescriptorSupport.copyOptionalText(specificGuidance, "specificGuidance");
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
