package dev.erst.gridgrind.excel;

import java.util.Locale;
import org.apache.poi.ss.util.CellReference;

/** Formats zero-based GridGrind row and column indexes using Excel-native equivalents. */
final class ExcelIndexDisplay {
  private ExcelIndexDisplay() {}

  static String excelRow(int zeroBasedIndex) {
    return "Excel row " + (zeroBasedIndex + 1);
  }

  static String excelColumn(int zeroBasedIndex) {
    return "Excel column " + CellReference.convertNumToColString(zeroBasedIndex);
  }

  static String rowValue(int zeroBasedIndex) {
    return zeroBasedIndex + " (" + excelRow(zeroBasedIndex) + ")";
  }

  static String columnValue(int zeroBasedIndex) {
    return zeroBasedIndex + " (" + excelColumn(zeroBasedIndex) + ")";
  }

  static String describe(String fieldName, int zeroBasedIndex) {
    return fieldName + " " + value(fieldName, zeroBasedIndex);
  }

  static String mustNotBeNegative(String fieldName, int value) {
    return fieldName
        + " "
        + value
        + " must not be negative; first addressable "
        + axis(fieldName)
        + " is "
        + firstCoordinate(fieldName);
  }

  static String mustNotBeLessThan(
      String lastFieldName, int lastValue, String firstFieldName, int firstValue) {
    return describe(lastFieldName, lastValue)
        + " must not be less than "
        + describe(firstFieldName, firstValue);
  }

  static String mustNotExceed(String fieldName, int actual, int maximum) {
    return describe(fieldName, actual) + " must not exceed " + value(fieldName, maximum);
  }

  static String value(String fieldName, int zeroBasedIndex) {
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
}
