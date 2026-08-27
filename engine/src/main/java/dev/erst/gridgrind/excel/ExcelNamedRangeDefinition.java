package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import java.util.Objects;

/** Immutable workbook-core definition of one named range to create or replace. */
public record ExcelNamedRangeDefinition(
    String name, ExcelNamedRangeScope scope, ExcelNamedRangeTarget target) {
  public ExcelNamedRangeDefinition {
    name = validateName(name);
    Objects.requireNonNull(scope, "scope must not be null");
    Objects.requireNonNull(target, "target must not be null");
  }

  /** Validates and canonicalizes one defined-name identifier. */
  public static String validateName(String name) {
    return ProtocolDefinedNameValidation.validateName(name);
  }

  /** Validates the minimum factual contract required to expose an observed workbook name. */
  static String validateObservedName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("observed name must not be blank");
    }
    return name;
  }
}
