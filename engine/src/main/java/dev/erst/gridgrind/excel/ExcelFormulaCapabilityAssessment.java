package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Engine-side assessment of one formula cell under the current evaluation environment. */
public record ExcelFormulaCapabilityAssessment(
    String sheetName,
    String address,
    String formula,
    ExcelFormulaCapabilityKind capability,
    ExcelFormulaCapabilityIssue issue,
    String message) {
  public ExcelFormulaCapabilityAssessment {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    Objects.requireNonNull(address, "address must not be null");
    if (address.isBlank()) {
      throw new IllegalArgumentException("address must not be blank");
    }
    Objects.requireNonNull(formula, "formula must not be null");
    if (formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
    Objects.requireNonNull(capability, "capability must not be null");
    if (capability == ExcelFormulaCapabilityKind.EVALUABLE_NOW && issue != null) {
      throw new IllegalArgumentException("issue is not permitted for EVALUABLE_NOW formulas");
    }
    if (capability != ExcelFormulaCapabilityKind.EVALUABLE_NOW && issue == null) {
      throw new IllegalArgumentException(
          "issue must be present for non-evaluable formula capabilities");
    }
    if (message != null && message.isBlank()) {
      throw new IllegalArgumentException("message must not be blank");
    }
  }
}
