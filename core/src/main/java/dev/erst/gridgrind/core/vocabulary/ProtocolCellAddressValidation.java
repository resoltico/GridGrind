package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.limits.Limit;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Protocol-owned validation helpers for single-cell A1-style addresses. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_ROWS_PER_WORKSHEET,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_COLUMNS_PER_WORKSHEET
})
public final class ProtocolCellAddressValidation {
  private static final int MAX_ROW_DIGITS = Integer.toString(Limit.MAX_ROW_INDEX + 1).length();

  private ProtocolCellAddressValidation() {}

  /** Validates one single-cell A1-style address and returns it unchanged. */
  public static String validateAddress(String address) {
    parseAddress(address);
    return address;
  }

  /** Parses a concrete position without retaining formula anchoring or spelling. */
  public static CellAddress parseAddress(String address) {
    Objects.requireNonNull(address, "address must not be null");
    if (address.isBlank()) {
      throw invalid(Violation.BLANK);
    }
    if (!address.matches("(?i)^\\$?[A-Z]{1,3}\\$?[1-9][0-9]*$")) {
      throw invalid(Violation.SYNTAX);
    }
    String normalized = address.replace("$", "").toUpperCase(Locale.ROOT);
    int split = 0;
    int column = 0;
    while (normalized.charAt(split) >= 'A' && normalized.charAt(split) <= 'Z') {
      column = column * 26 + normalized.charAt(split) - 'A' + 1;
      split++;
    }
    String digits = normalized.substring(split);
    if (digits.length() > MAX_ROW_DIGITS) {
      throw invalid(Violation.BOUNDS);
    }
    return new CellAddress(Integer.parseInt(digits) - 1, column - 1);
  }

  /** Returns the first address rule violation, if any. */
  public static Optional<Violation> violation(String address) {
    try {
      parseAddress(address);
      return Optional.empty();
    } catch (dev.erst.gridgrind.core.problem.InvalidRequestException failure) {
      var problem =
          (dev.erst.gridgrind.core.problem.FieldValidationProblem) failure.requestProblem();
      return Optional.of(
          switch (problem.rule()) {
            case dev.erst.gridgrind.core.problem.FieldValidationBasicRule.NON_BLANK ->
                Violation.BLANK;
            case dev.erst.gridgrind.core.problem.FieldValidationAddressRule.ADDRESS_SYNTAX ->
                Violation.SYNTAX;
            default -> Violation.BOUNDS;
          });
    }
  }

  static dev.erst.gridgrind.core.problem.InvalidRequestException invalid(Violation violation) {
    var mapped =
        dev.erst.gridgrind.core.problem.FieldValidationProblemMappers.address("address", violation);
    return new dev.erst.gridgrind.core.problem.InvalidRequestException(
        new dev.erst.gridgrind.core.problem.FieldValidationProblem(
            mapped.fieldName(), Optional.empty(), mapped.rule(), mapped.operands()),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        null);
  }

  /** Structured address violation family. */
  public enum Violation {
    BLANK,
    SYNTAX,
    BOUNDS
  }
}
