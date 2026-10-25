package dev.erst.gridgrind.contract.action;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.json.FieldValidationBasicRule;
import dev.erst.gridgrind.contract.json.FieldValidationBoundRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/** Collection-shaped validation for mutation actions. */
final class MutationActionCollectionValidation {
  private MutationActionCollectionValidation() {}

  static List<List<CellInput>> copyRows(List<List<CellInput>> rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    List<List<CellInput>> copy = new ArrayList<>(rows.size());
    for (List<CellInput> row : rows) {
      copy.add(new ArrayList<>(Objects.requireNonNull(row, "rows must not contain null rows")));
    }
    return java.util.Collections.unmodifiableList(copy);
  }

  static List<List<CellInput>> freezeRows(List<List<CellInput>> rows) {
    return rows.stream().map(List::copyOf).toList();
  }

  static List<String> copySheetNames(List<String> sheetNames, String fieldName) {
    Objects.requireNonNull(sheetNames, fieldName + " must not be null");
    List<String> copy = new ArrayList<>(sheetNames);
    for (int index = 0; index < copy.size(); index++) {
      MutationActionNameValidation.requireSheetName(copy.get(index), fieldName + "[" + index + "]");
    }
    return List.copyOf(copy);
  }

  static void requireDistinct(List<String> values, String fieldName) {
    if (new LinkedHashSet<>(values).size() != values.size()) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
    }
  }

  static void requireRectangularRows(List<List<CellInput>> rows) {
    if (rows.isEmpty()) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField("rows", FieldValidationBasicRule.ROWS_NON_EMPTY));
    }
    int expectedWidth = -1;
    for (List<CellInput> row : rows) {
      Objects.requireNonNull(row, "rows must not contain null rows");
      if (row.isEmpty()) {
        throw MutationActionNameValidation.invalidField(
            FieldValidationProblem.atField("rows", FieldValidationBasicRule.EMPTY_ROWS));
      }
      if (expectedWidth < 0) {
        expectedWidth = row.size();
      } else if (row.size() != expectedWidth) {
        throw MutationActionNameValidation.invalidField(
            FieldValidationProblem.atField("rows", FieldValidationBoundRule.RECTANGULAR_MATRIX));
      }
      for (CellInput value : row) {
        Objects.requireNonNull(value, "rows must not contain null cell values");
      }
    }
  }
}
