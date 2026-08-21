package dev.erst.gridgrind.contract.step;

/** One precise static request-contract violation ready for diagnostic projection. */
public record WorkbookStaticViolation(String jsonPath, String message) {
  public WorkbookStaticViolation {
    if (jsonPath == null || jsonPath.isBlank()) {
      throw new IllegalArgumentException("jsonPath must not be blank");
    }
    if (message == null || message.isBlank()) {
      throw new IllegalArgumentException("message must not be blank");
    }
  }
}
