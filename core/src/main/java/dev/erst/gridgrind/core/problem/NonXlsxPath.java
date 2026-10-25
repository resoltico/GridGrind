package dev.erst.gridgrind.core.problem;

import java.util.Optional;

/** One workbook path did not end in the required `.xlsx` extension. */
public record NonXlsxPath(String actualExtension, Optional<String> jsonPathValue)
    implements RequestProblemDescriptor.Invariant {
  public NonXlsxPath {
    actualExtension =
        RequestProblemDescriptorSupport.requireNonBlank(actualExtension, "actualExtension");
    jsonPathValue = RequestProblemDescriptorSupport.copyJsonPath(jsonPathValue);
  }

  @Override
  public Optional<String> jsonPath() {
    return jsonPathValue;
  }
}
