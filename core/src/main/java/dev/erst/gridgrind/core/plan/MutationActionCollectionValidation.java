package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.problem.FieldValidationBasicRule;
import dev.erst.gridgrind.core.problem.FieldValidationBoundRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/** Collection-shaped validation for mutation actions. */
final class MutationActionCollectionValidation {
  private MutationActionCollectionValidation() {}

  static List<List<CellValueSpec>> copyRows(List<List<CellValueSpec>> rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    List<List<CellValueSpec>> copy = new ArrayList<>(rows.size());
    for (List<CellValueSpec> row : rows) {
      copy.add(new ArrayList<>(Objects.requireNonNull(row, "rows must not contain null rows")));
    }
    return java.util.Collections.unmodifiableList(copy);
  }

  static List<List<CellValueSpec>> freezeRows(List<List<CellValueSpec>> rows) {
    return rows.stream().map(List::copyOf).toList();
  }

  static List<SheetName> copySheetNames(List<SheetName> sheetNames, String fieldName) {
    Objects.requireNonNull(sheetNames, fieldName + " must not be null");
    List<SheetName> copy = new ArrayList<>(sheetNames);
    for (int index = 0; index < copy.size(); index++) {
      Objects.requireNonNull(copy.get(index), fieldName + "[" + index + "] must not be null");
    }
    return List.copyOf(copy);
  }

  static void requireDistinct(List<String> values, String fieldName) {
    if (new LinkedHashSet<>(values).size() != values.size()) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
    }
  }

  static void requireRectangularRows(List<List<CellValueSpec>> rows) {
    if (rows.isEmpty()) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField("rows", FieldValidationBasicRule.ROWS_NON_EMPTY));
    }
    int expectedWidth = -1;
    for (List<CellValueSpec> row : rows) {
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
      for (CellValueSpec value : row) {
        Objects.requireNonNull(value, "rows must not contain null cell values");
      }
    }
  }
}
