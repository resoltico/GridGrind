package dev.erst.gridgrind.contract.json;

/** Address and range rules tied to Excel A1-style coordinates. */
public enum FieldValidationAddressRule implements FieldValidationRule {
  ADDRESS_SYNTAX,
  ADDRESS_BOUNDS,
  RANGE_RECTANGULAR_SYNTAX;

  @Override
  public String message(FieldValidationProblem problem) {
    return switch (this) {
      case ADDRESS_SYNTAX -> problem.fieldName() + " must be a single-cell A1-style address";
      case ADDRESS_BOUNDS -> problem.fieldName() + " must be within Excel .xlsx bounds";
      case RANGE_RECTANGULAR_SYNTAX ->
          problem.fieldName() + " must be a rectangular A1-style range with at most one ':'";
    };
  }

  @Override
  public String resolution(FieldValidationProblem problem) {
    return switch (this) {
      case ADDRESS_SYNTAX, ADDRESS_BOUNDS ->
          "Use a single-cell A1-style address such as A1 or BC12 within Excel .xlsx bounds for"
              + " field '"
              + problem.fieldName()
              + "'.";
      case RANGE_RECTANGULAR_SYNTAX ->
          "Provide a rectangular A1-style range with at most one ':' for field '"
              + problem.fieldName()
              + "'.";
    };
  }
}
