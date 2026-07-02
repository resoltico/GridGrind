package dev.erst.gridgrind.contract.selector;

import dev.erst.gridgrind.contract.json.FieldValidationBasicRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Shared distinct-list validation for selector families. */
final class SelectorListValidation {
  private SelectorListValidation() {}

  static List<String> copyDistinctAddresses(List<String> addresses, String fieldName) {
    Objects.requireNonNull(addresses, fieldName + " must not be null");
    return copyDistinctStrings(addresses, fieldName, false, true, false);
  }

  static List<String> copyDistinctRanges(List<String> ranges, String fieldName) {
    Objects.requireNonNull(ranges, fieldName + " must not be null");
    return copyDistinctStrings(ranges, fieldName, false, false, true);
  }

  static List<String> copyDistinctSheetNames(List<String> sheetNames, String fieldName) {
    Objects.requireNonNull(sheetNames, fieldName + " must not be null");
    return copyDistinctStrings(sheetNames, fieldName, true, false, false);
  }

  static List<String> copyDistinctDefinedNames(List<String> names, String fieldName) {
    Objects.requireNonNull(names, fieldName + " must not be null");
    return copyDistinctStrings(names, fieldName, false, false, false);
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
          "NAMED_RANGE_BY_NAME|" + byName.name().toUpperCase(Locale.ROOT);
      case NamedRangeSelector.WorkbookScope workbookScope ->
          "NAMED_RANGE_WORKBOOK_SCOPE|" + workbookScope.name().toUpperCase(Locale.ROOT);
      case NamedRangeSelector.SheetScope sheetScope ->
          "NAMED_RANGE_SHEET_SCOPE|"
              + sheetScope.sheetName().toUpperCase(Locale.ROOT)
              + "|"
              + sheetScope.name().toUpperCase(Locale.ROOT);
    };
  }

  private static List<String> copyDistinctStrings(
      List<String> values,
      String fieldName,
      boolean sheetNames,
      boolean addresses,
      boolean ranges) {
    List<String> copy = new ArrayList<>(values);
    if (copy.isEmpty()) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_EMPTY));
    }
    if (addresses) {
      SelectorNumberValidation.requireReadCellCount(copy.size(), fieldName);
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      String indexedFieldName = fieldName + "[" + index + "]";
      String validated =
          sheetNames
              ? SelectorValueValidation.requireSheetName(copy.get(index), indexedFieldName)
              : addresses
                  ? SelectorValueValidation.requireAddress(copy.get(index), indexedFieldName)
                  : ranges
                      ? SelectorValueValidation.requireRange(copy.get(index), indexedFieldName)
                      : SelectorValueValidation.requireDefinedName(
                          copy.get(index), indexedFieldName);
      String key = validated.toUpperCase(Locale.ROOT);
      if (!unique.add(key)) {
        throw SelectorTextValidation.invalidField(
            FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.DUPLICATES));
      }
    }
    return List.copyOf(copy);
  }
}
