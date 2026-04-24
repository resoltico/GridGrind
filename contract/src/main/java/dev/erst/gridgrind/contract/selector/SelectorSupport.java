package dev.erst.gridgrind.contract.selector;

import dev.erst.gridgrind.contract.dto.ProtocolCellAddressValidation;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.excel.foundation.ExcelAddressLists;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.foundation.ExcelReadLimits;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Shared validation helpers for selector families. */
final class SelectorSupport {
  private SelectorSupport() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static String requireSheetName(String value, String fieldName) {
    String validated = requireNonBlank(value, fieldName);
    ExcelSheetNames.requireValid(validated, fieldName);
    return validated;
  }

  static String requireDefinedName(String value, String fieldName) {
    try {
      return ProtocolDefinedNameValidation.validateName(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(
          prefixedValidationMessage(fieldName, exception.getMessage()), exception);
    }
  }

  static String requirePivotTableName(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    try {
      return ExcelPivotTableNaming.validateName(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(
          prefixedValidationMessage(fieldName, exception.getMessage()), exception);
    }
  }

  static String requireAddress(String value, String fieldName) {
    try {
      return ProtocolCellAddressValidation.validateAddress(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(
          prefixedValidationMessage(fieldName, exception.getMessage()), exception);
    }
  }

  static String requireRange(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    String[] parts = value.split(":", -1);
    if (parts.length > 2) {
      throw new IllegalArgumentException(
          fieldName + " must be a rectangular A1-style range with at most one ':'");
    }
    requireAddress(parts[0], fieldName);
    if (parts.length == 2) {
      requireAddress(parts[1], fieldName);
    }
    return value;
  }

  static int requirePositive(int value, String fieldName) {
    if (value <= 0) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
    return value;
  }

  static int requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
    return value;
  }

  static int requireNonZero(int value, String fieldName) {
    if (value == 0) {
      throw new IllegalArgumentException(fieldName + " must not be 0");
    }
    return value;
  }

  static int requireRowIndexWithinBounds(int value, String fieldName) {
    requireNonNegative(value, fieldName);
    if (value > ExcelRowSpan.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(fieldName + " must be within Excel .xlsx row bounds");
    }
    return value;
  }

  static int requireColumnIndexWithinBounds(int value, String fieldName) {
    requireNonNegative(value, fieldName);
    if (value > ExcelColumnSpan.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(fieldName + " must be within Excel .xlsx column bounds");
    }
    return value;
  }

  static void requireWindowSize(int rowCount, int columnCount) {
    long cells = (long) rowCount * columnCount;
    if (cells > ExcelReadLimits.MAX_WINDOW_CELLS) {
      throw new IllegalArgumentException(
          "rowCount * columnCount must not exceed "
              + ExcelReadLimits.MAX_WINDOW_CELLS
              + " but was "
              + cells);
    }
  }

  static List<String> copyDistinctAddresses(List<String> addresses, String fieldName) {
    Objects.requireNonNull(addresses, fieldName + " must not be null");
    if ("addresses".equals(fieldName)) {
      return ExcelAddressLists.copyNonEmptyDistinctAddresses(addresses);
    }
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
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      String name = copy.get(index);
      String validated;
      try {
        validated = requirePivotTableName(name, "name");
      } catch (IllegalArgumentException exception) {
        throw new IllegalArgumentException(
            fieldName + "[" + index + "] " + exception.getMessage(), exception);
      }
      if (!unique.add(validated.toUpperCase(Locale.ROOT))) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }

  static <T> List<T> copyDistinctValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    Set<T> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      T value = copy.get(index);
      Objects.requireNonNull(value, fieldName + "[" + index + "] must not be null");
      if (!unique.add(value)) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }

  static List<NamedRangeSelector.Ref> copyDistinctNamedRangeRefs(
      List<NamedRangeSelector.Ref> selectors, String fieldName) {
    Objects.requireNonNull(selectors, fieldName + " must not be null");
    List<NamedRangeSelector.Ref> copy = new ArrayList<>(selectors);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      NamedRangeSelector.Ref selector = copy.get(index);
      Objects.requireNonNull(selector, fieldName + "[" + index + "] must not be null");
      String key = namedRangeRefIdentity(selector);
      if (!unique.add(key)) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }

  static String absoluteA1Address(int rowIndex, int columnIndex) {
    return columnLabel(columnIndex) + (rowIndex + 1);
  }

  static int columnIndex(String address) {
    String normalized = address.replace("$", "").toUpperCase(Locale.ROOT);
    int splitIndex = 0;
    while (splitIndex < normalized.length() && Character.isLetter(normalized.charAt(splitIndex))) {
      splitIndex++;
    }
    int result = 0;
    for (int index = 0; index < splitIndex; index++) {
      result = (result * 26) + (normalized.charAt(index) - 'A' + 1);
    }
    return result - 1;
  }

  static int rowIndex(String address) {
    String normalized = address.replace("$", "");
    int splitIndex = 0;
    while (splitIndex < normalized.length() && !Character.isDigit(normalized.charAt(splitIndex))) {
      splitIndex++;
    }
    return Integer.parseInt(normalized.substring(splitIndex)) - 1;
  }

  private static String columnLabel(int columnIndex) {
    int value = columnIndex + 1;
    StringBuilder builder = new StringBuilder();
    while (value > 0) {
      int remainder = (value - 1) % 26;
      builder.append((char) ('A' + remainder));
      value = (value - 1) / 26;
    }
    return builder.reverse().toString();
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

  static String prefixedValidationMessage(String fieldName, String message) {
    if (message == null || message.isBlank() || message.startsWith(fieldName + " ")) {
      return message;
    }
    return fieldName + " " + message;
  }

  private static List<String> copyDistinctStrings(
      List<String> values,
      String fieldName,
      boolean sheetNames,
      boolean addresses,
      boolean ranges) {
    List<String> copy = new ArrayList<>(values);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      String value = copy.get(index);
      String validated;
      try {
        validated =
            sheetNames
                ? requireSheetName(value, "sheetName")
                : addresses
                    ? requireAddress(value, "address")
                    : ranges ? requireRange(value, "range") : requireDefinedName(value, "name");
      } catch (IllegalArgumentException exception) {
        throw new IllegalArgumentException(
            fieldName + "[" + index + "] " + exception.getMessage(), exception);
      }
      String key = validated.toUpperCase(Locale.ROOT);
      if (!unique.add(key)) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
      }
    }
    return List.copyOf(copy);
  }
}
