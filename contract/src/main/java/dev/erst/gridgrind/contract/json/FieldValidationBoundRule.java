package dev.erst.gridgrind.contract.json;

/** Index, span, and matrix rules whose wording depends on one bounded operand. */
public enum FieldValidationBoundRule implements FieldValidationRule {
  ROW_INDEX_BOUNDS,
  COLUMN_INDEX_BOUNDS,
  COLLECTION_SIZE_LIMIT,
  ORDERED_SPAN,
  RECTANGULAR_MATRIX,
  WINDOW_SIZE_PRODUCT;

  @Override
  public String message(FieldValidationProblem problem) {
    return switch (this) {
      case ROW_INDEX_BOUNDS ->
          problem.fieldName() + " must not exceed " + problem.operand(0) + " (Excel row limit)";
      case COLUMN_INDEX_BOUNDS ->
          problem.fieldName() + " must not exceed " + problem.operand(0) + " (Excel column limit)";
      case COLLECTION_SIZE_LIMIT ->
          problem.fieldName()
              + " must not exceed "
              + problem.operand(0)
              + " but was "
              + problem.operand(1);
      case ORDERED_SPAN -> problem.fieldName() + " must not be less than " + problem.operand(0);
      case RECTANGULAR_MATRIX -> problem.fieldName() + " must describe a rectangular matrix";
      case WINDOW_SIZE_PRODUCT ->
          problem.fieldName()
              + " must not exceed "
              + problem.operand(0)
              + " but was "
              + problem.operand(1);
    };
  }

  @Override
  public String resolution(FieldValidationProblem problem) {
    return switch (this) {
      case ROW_INDEX_BOUNDS ->
          "Provide a row index between 0 and "
              + problem.operand(0)
              + " for field '"
              + problem.fieldName()
              + "'.";
      case COLUMN_INDEX_BOUNDS ->
          "Provide a column index between 0 and "
              + problem.operand(0)
              + " for field '"
              + problem.fieldName()
              + "'.";
      case COLLECTION_SIZE_LIMIT ->
          "Reduce field '"
              + problem.fieldName()
              + "' to at most "
              + problem.operand(0)
              + " values.";
      case ORDERED_SPAN ->
          "Ensure '"
              + problem.fieldName()
              + "' is greater than or equal to '"
              + problem.operand(0)
              + "'.";
      case RECTANGULAR_MATRIX ->
          "Ensure field '"
              + problem.fieldName()
              + "' describes a rectangular matrix with the same width in every row.";
      case WINDOW_SIZE_PRODUCT ->
          "Reduce rowCount and columnCount so their product does not exceed "
              + problem.operand(0)
              + ".";
    };
  }
}
