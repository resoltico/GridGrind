package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.problem.FieldValidationAddressRule;
import dev.erst.gridgrind.core.problem.FieldValidationBasicRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.FieldValidationProblemMappers;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import java.util.Objects;
import java.util.Optional;

/** One local A1 cell or rectangular range, preserving authored endpoint order and spelling. */
public record RangeRef(String value) implements WireString {
  public RangeRef {
    Objects.requireNonNull(value, "range must not be null");
    if (value.isBlank()) {
      throw invalid(FieldValidationProblem.detached("range", FieldValidationBasicRule.NON_BLANK));
    }
    String[] endpoints = value.split(":", -1);
    if (endpoints.length > 2) {
      throw invalid(
          FieldValidationProblem.detached(
              "range", FieldValidationAddressRule.RANGE_RECTANGULAR_SYNTAX));
    }
    for (String endpoint : endpoints) {
      ProtocolCellAddressValidation.violation(endpoint)
          .ifPresent(
              violation -> {
                throw invalid(FieldValidationProblemMappers.address("range", violation));
              });
    }
  }

  private static InvalidRequestException invalid(FieldValidationProblem problem) {
    return new InvalidRequestException(
        new FieldValidationProblem(
            problem.fieldName(), Optional.empty(), problem.rule(), problem.operands()),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        null);
  }

  @Override
  public String toString() {
    return value;
  }
}
