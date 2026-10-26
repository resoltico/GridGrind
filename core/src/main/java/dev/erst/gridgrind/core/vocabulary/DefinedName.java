package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.FieldValidationProblemMappers;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import java.util.Objects;
import java.util.Optional;

/** A validated Excel defined-name identifier. */
public record DefinedName(String value) implements WireString {
  public DefinedName {
    Objects.requireNonNull(value, "name must not be null");
    ProtocolDefinedNameValidation.violation(value)
        .ifPresent(
            violation -> {
              var mapped = FieldValidationProblemMappers.definedName("name", violation);
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
