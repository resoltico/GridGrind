package dev.erst.gridgrind.core.problem;

/** Authoritative wording for unsupported authored civil dates and mapped date representations. */
public enum FieldValidationTemporalRule implements FieldValidationRule {
  DATE_RANGE,
  MILLISECOND_PRECISION,
  ACTIVE_EPOCH,
  XML_DATE_LEXICAL,
  XML_DATE_OFFSET;

  @Override
  public String message(FieldValidationProblem problem) {
    return switch (this) {
      case DATE_RANGE ->
          problem.fieldName()
              + " must be between "
              + problem.operand(0)
              + " and "
              + problem.operand(1)
              + ".";
      case MILLISECOND_PRECISION ->
          problem.fieldName()
              + " supports millisecond precision; finer fractional seconds are not supported.";
      case ACTIVE_EPOCH ->
          problem.fieldName()
              + " must be on or after "
              + problem.operand(0)
              + " for the "
              + problem.operand(1)
              + " workbook date system.";
      case XML_DATE_LEXICAL -> problem.fieldName() + " must be a valid XML Schema date.";
      case XML_DATE_OFFSET ->
          problem.fieldName()
              + " must be timezone-free; an XML Schema date offset cannot be retained in an Excel"
              + " civil date.";
    };
  }

  @Override
  public String resolution(FieldValidationProblem problem) {
    return switch (this) {
      case DATE_RANGE, ACTIVE_EPOCH ->
          "Provide a civil date within the receiving workbook's supported calendar range.";
      case MILLISECOND_PRECISION -> "Provide a DATE_TIME value with at most millisecond precision.";
      case XML_DATE_LEXICAL -> "Provide a valid Gregorian XML Schema date.";
      case XML_DATE_OFFSET ->
          "Provide an explicitly timezone-free civil date rather than discarding its offset"
              + " implicitly.";
    };
  }
}
