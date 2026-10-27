package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.FieldValidationProblemMappers;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import java.util.Objects;
import java.util.Optional;

/** One local A1-style cell address within Excel .xlsx bounds. */
public record CellAddress(String value) implements WireString {
  public CellAddress {
    Objects.requireNonNull(value, "address must not be null");
    ProtocolCellAddressValidation.violation(value)
        .ifPresent(
            violation -> {
              FieldValidationProblem mapped =
                  FieldValidationProblemMappers.address("address", violation);
              throw new InvalidRequestException(
                  new FieldValidationProblem(
                      mapped.fieldName(), Optional.empty(), mapped.rule(), mapped.operands()),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty(),
                  null);
            });
  }

  @Override
  public String toString() {
    return value;
  }
}
