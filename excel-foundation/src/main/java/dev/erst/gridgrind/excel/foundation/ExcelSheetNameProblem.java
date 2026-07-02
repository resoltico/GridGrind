package dev.erst.gridgrind.excel.foundation;

import java.util.List;

/** One typed Excel sheet-name validation failure with stable operands. */
public record ExcelSheetNameProblem(ExcelSheetNameRule rule, List<String> operands) {
  /** Validates one typed sheet-name problem and defensively copies its operands. */
  public ExcelSheetNameProblem {
    java.util.Objects.requireNonNull(rule, "rule must not be null");
    operands = List.copyOf(java.util.Objects.requireNonNull(operands, "operands must not be null"));
  }

  /** Returns the blank-sheet-name problem. */
  public static ExcelSheetNameProblem blank() {
    return new ExcelSheetNameProblem(ExcelSheetNameRule.BLANK, List.of());
  }

  /** Returns the overlength-sheet-name problem carrying the authored value. */
  public static ExcelSheetNameProblem tooLong(String value) {
    return new ExcelSheetNameProblem(ExcelSheetNameRule.TOO_LONG, List.of(value));
  }

  /** Returns the boundary-quote problem carrying the authored value. */
  public static ExcelSheetNameProblem boundaryQuote(String value) {
    return new ExcelSheetNameProblem(ExcelSheetNameRule.BOUNDARY_QUOTE, List.of(value));
  }

  /** Returns the invalid-character problem carrying display text, position, and authored value. */
  public static ExcelSheetNameProblem invalidCharacter(
      String displayCharacter, int position, String value) {
    return new ExcelSheetNameProblem(
        ExcelSheetNameRule.INVALID_CHARACTER,
        List.of(displayCharacter, Integer.toString(position), value));
  }

  /** Renders the canonical public message for this sheet-name problem. */
  public String message(String fieldName) {
    return switch (rule) {
      case BLANK -> fieldName + " must not be blank";
      case TOO_LONG -> fieldName + " must not exceed 31 characters: " + operand(0);
      case BOUNDARY_QUOTE ->
          fieldName + " must not begin or end with a single quote: " + operand(0);
      case INVALID_CHARACTER ->
          fieldName
              + " contains invalid Excel character "
              + operand(0)
              + " at position "
              + operand(1)
              + ": "
              + operand(2);
    };
  }

  /** Returns one stable operand by index for downstream mappers and renderers. */
  public String operand(int index) {
    return operands.get(index);
  }
}
