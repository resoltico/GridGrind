package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.problem.FieldValidationBasicRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.core.vocabulary.WireString;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Shared distinct-list validation for selector families. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_GET_CELLS_ADDRESS_LIST
})
final class SelectorListValidation {
  private SelectorListValidation() {}

  static List<CellAddress> copyDistinctAddresses(List<CellAddress> addresses, String fieldName) {
    Objects.requireNonNull(addresses, fieldName + " must not be null");
    SelectorNumberValidation.requireReadCellCount(addresses.size(), fieldName);
    return copyDistinctCaseInsensitive(addresses, fieldName);
  }

  static List<RangeRef> copyDistinctRanges(List<RangeRef> ranges, String fieldName) {
    Objects.requireNonNull(ranges, fieldName + " must not be null");
    List<RangeRef> copy = new ArrayList<>(ranges);
    if (copy.isEmpty()) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_EMPTY));
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      RangeRef range =
          SelectorValueValidation.requireOrderedRange(
              copy.get(index), fieldName + "[" + index + "]");
      if (!unique.add(range.value().toUpperCase(Locale.ROOT))) {
        throw SelectorTextValidation.invalidField(
            FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
      }
    }
    return List.copyOf(copy);
  }

  static List<SheetName> copyDistinctSheetNames(List<SheetName> sheetNames, String fieldName) {
    return copyDistinctCaseInsensitive(sheetNames, fieldName);
  }

  private static <T extends WireString> List<T> copyDistinctCaseInsensitive(
      List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    if (values.isEmpty()) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_EMPTY));
    }
    Set<String> unique = new LinkedHashSet<>();
    for (T value : values) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
      if (!unique.add(value.value().toUpperCase(Locale.ROOT))) {
        throw SelectorTextValidation.invalidField(
            FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
      }
    }
    return List.copyOf(values);
  }

  static List<DefinedName> copyDistinctNamedRangeNames(List<DefinedName> names, String fieldName) {
    return copyDistinctCaseInsensitive(names, fieldName);
  }

  static List<TableName> copyDistinctTableNames(List<TableName> names, String fieldName) {
    return copyDistinctCaseInsensitive(names, fieldName);
  }

  static List<String> copyDistinctPivotTableNames(List<String> names, String fieldName) {
    Objects.requireNonNull(names, fieldName + " must not be null");
    List<String> copy = new ArrayList<>(names);
    if (copy.isEmpty()) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_EMPTY));
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      String validated =
          SelectorValueValidation.requirePivotTableName(
              copy.get(index), fieldName + "[" + index + "]");
      if (!unique.add(validated.toUpperCase(Locale.ROOT))) {
        throw SelectorTextValidation.invalidField(
            FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
      }
    }
    return List.copyOf(copy);
  }

  static <T> List<T> copyDistinctValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values);
    if (copy.isEmpty()) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_EMPTY));
    }
    Set<T> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      T value = copy.get(index);
      Objects.requireNonNull(value, fieldName + "[" + index + "] must not be null");
      if (!unique.add(value)) {
        throw SelectorTextValidation.invalidField(
            FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
      }
    }
    return List.copyOf(copy);
  }

  static List<NamedRangeSelector.Ref> copyDistinctNamedRangeRefs(
      List<NamedRangeSelector.Ref> selectors, String fieldName) {
    Objects.requireNonNull(selectors, fieldName + " must not be null");
    List<NamedRangeSelector.Ref> copy = new ArrayList<>(selectors);
    if (copy.isEmpty()) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_EMPTY));
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      NamedRangeSelector.Ref selector = copy.get(index);
      Objects.requireNonNull(selector, fieldName + "[" + index + "] must not be null");
      String key = namedRangeRefIdentity(selector);
      if (!unique.add(key)) {
        throw SelectorTextValidation.invalidField(
            FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
      }
    }
    return List.copyOf(copy);
  }

  private static String namedRangeRefIdentity(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName ->
          "NAMED_RANGE_BY_NAME|" + byName.name().value().toUpperCase(Locale.ROOT);
      case NamedRangeSelector.WorkbookScope workbookScope ->
          "NAMED_RANGE_WORKBOOK_SCOPE|" + workbookScope.name().value().toUpperCase(Locale.ROOT);
      case NamedRangeSelector.SheetScope sheetScope ->
          "NAMED_RANGE_SHEET_SCOPE|"
              + sheetScope.sheetName().value().toUpperCase(Locale.ROOT)
              + "|"
              + sheetScope.name().value().toUpperCase(Locale.ROOT);
    };
  }
}
