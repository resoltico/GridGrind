package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reports JSON grammar that cannot be assigned a GridGrind request shape. */
record RequestInvalidJson(
    String message,
    Optional<String> affectedJsonPath,
    Optional<Long> byteOffset,
    Optional<Integer> jsonLine,
    Optional<Integer> jsonColumn)
    implements RequestStructuralProblem {
  RequestInvalidJson {
    message = RequestStructuralProblemSupport.requireText(message, "message");
    affectedJsonPath = RequestStructuralProblemSupport.copyJsonPath(affectedJsonPath);
    byteOffset = RequestStructuralProblemSupport.copyByteOffset(byteOffset);
    jsonLine = jsonLine.map(value -> value < 1 ? null : value);
    jsonColumn = jsonColumn.map(value -> value < 1 ? null : value);
    if (jsonLine.isPresent() != jsonColumn.isPresent()) {
      throw new IllegalArgumentException("jsonLine and jsonColumn must be supplied together");
    }
  }

  RequestInvalidJson(String message, Optional<String> affectedJsonPath, Optional<Long> byteOffset) {
    this(message, affectedJsonPath, byteOffset, Optional.empty(), Optional.empty());
  }

  @Override
  public Optional<String> jsonPath() {
    return affectedJsonPath;
  }
}
