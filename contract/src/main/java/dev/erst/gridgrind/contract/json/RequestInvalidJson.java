package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reports JSON grammar that cannot be assigned a GridGrind request shape. */
record RequestInvalidJson(
    String message, Optional<String> affectedJsonPath, Optional<Long> byteOffset)
    implements RequestStructuralProblem {
  RequestInvalidJson {
    message = RequestStructuralProblemSupport.requireText(message, "message");
    affectedJsonPath = RequestStructuralProblemSupport.copyJsonPath(affectedJsonPath);
    byteOffset = RequestStructuralProblemSupport.copyByteOffset(byteOffset);
  }

  @Override
  public Optional<String> jsonPath() {
    return affectedJsonPath;
  }
}
