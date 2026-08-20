package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reports one authored number that an IEEE-backed request field cannot retain exactly. */
record RequestNumberNotRepresentable(String path, String token, long tokenByteOffset)
    implements RequestStructuralProblem {
  RequestNumberNotRepresentable {
    path = RequestStructuralProblemSupport.requireText(path, "path");
    token = RequestStructuralProblemSupport.requireText(token, "token");
    RequestStructuralProblemSupport.requireByteOffset(tokenByteOffset);
  }

  @Override
  public Optional<String> jsonPath() {
    return Optional.of(path);
  }

  @Override
  public Optional<Long> byteOffset() {
    return Optional.of(tokenByteOffset);
  }

  @Override
  public String message() {
    return "Field '"
        + path
        + "' number '"
        + token
        + "' cannot be represented exactly; store identifiers or precision-sensitive values as TEXT.";
  }
}
