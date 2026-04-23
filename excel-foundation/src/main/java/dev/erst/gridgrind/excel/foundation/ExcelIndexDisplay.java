package dev.erst.gridgrind.excel.foundation;

import java.util.Locale;

/** Formats zero-based GridGrind row and column indexes using Excel-native equivalents. */
public final class ExcelIndexDisplay {
  private ExcelIndexDisplay() {}

  /** Formats one zero-based row index as an Excel-facing row label. */
  public static String excelRow(int zeroBasedIndex) {
    return "Excel row " + (zeroBasedIndex + 1);
  }

  /** Formats one zero-based column index as an Excel-facing column label. */
  public static String excelColumn(int zeroBasedIndex) {
    return "Excel column " + excelColumnLabel(zeroBasedIndex);
  }

  /** Returns the raw zero-based row index plus its Excel-facing row label. */
  public static String rowValue(int zeroBasedIndex) {
    return zeroBasedIndex + " (" + excelRow(zeroBasedIndex) + ")";
  }

  /** Returns the raw zero-based column index plus its Excel-facing column label. */
  public static String columnValue(int zeroBasedIndex) {
    return zeroBasedIndex + " (" + excelColumn(zeroBasedIndex) + ")";
  }

  /** Describes one row or column field using its human-readable Excel coordinate. */
  public static String describe(String fieldName, int zeroBasedIndex) {
    return fieldName + " " + value(fieldName, zeroBasedIndex);
  }

  /** Builds the standard non-negative-validation message for one row or column field. */
  public static String mustNotBeNegative(String fieldName, int value) {
    return fieldName
        + " "
        + value
        + " must not be negative; first addressable "
        + axis(fieldName)
        + " is "
        + firstCoordinate(fieldName);
  }

  /** Builds the standard ordering-validation message for paired row or column fields. */
  public static String mustNotBeLessThan(
      String lastFieldName, int lastValue, String firstFieldName, int firstValue) {
    return describe(lastFieldName, lastValue)
        + " must not be less than "
        + describe(firstFieldName, firstValue);
  }

  /** Builds the standard maximum-boundary message for one row or column field. */
  public static String mustNotExceed(String fieldName, int actual, int maximum) {
    return describe(fieldName, actual) + " must not exceed " + value(fieldName, maximum);
  }

  /** Formats one row or column field value using the correct Excel coordinate family. */
  public static String value(String fieldName, int zeroBasedIndex) {
    return isColumnField(fieldName) ? columnValue(zeroBasedIndex) : rowValue(zeroBasedIndex);
  }

  private static boolean isColumnField(String fieldName) {
    return fieldName.toLowerCase(Locale.ROOT).contains("column");
  }

  private static String axis(String fieldName) {
    return isColumnField(fieldName) ? "column" : "row";
  }

  private static String firstCoordinate(String fieldName) {
    return isColumnField(fieldName) ? excelColumn(0) : excelRow(0);
  }

  private static String excelColumnLabel(int zeroBasedIndex) {
    int columnIndex = zeroBasedIndex;
    StringBuilder label = new StringBuilder();
    do {
      int remainder = columnIndex % 26;
      label.append((char) ('A' + remainder));
      columnIndex = (columnIndex / 26) - 1;
    } while (columnIndex >= 0);
    return label.reverse().toString();
  }
}
