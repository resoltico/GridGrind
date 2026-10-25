package dev.erst.gridgrind.excel.foundation;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

/** OOXML-storable cell error literals that GridGrind may author directly into workbook cells. */
public enum ExcelStoredCellErrorLiteral {
  NULL("#NULL!"),
  DIV0("#DIV/0!"),
  VALUE("#VALUE!"),
  REF("#REF!"),
  NAME("#NAME?"),
  NUM("#NUM!"),
  NA("#N/A");

  private static final Map<String, ExcelStoredCellErrorLiteral> BY_WIRE_VALUE =
      Arrays.stream(values())
          .collect(
              java.util.stream.Collectors.toUnmodifiableMap(
                  ExcelStoredCellErrorLiteral::wireValue, literal -> literal));

  private final String wireValue;

  ExcelStoredCellErrorLiteral(String wireValue) {
    this.wireValue = wireValue;
  }

  /** Returns the published wire token for this stored error literal. */
  public String wireValue() {
    return wireValue;
  }

  /** Resolves one published stored-cell error token to its owned literal. */
  public static ExcelStoredCellErrorLiteral fromWireValue(String wireValue) {
    if (wireValue == null) {
      throw new IllegalArgumentException("wireValue must not be null");
    }
    if (!isSupportedWireValue(wireValue)) {
      throw new IllegalArgumentException(
          "Unsupported stored Excel cell error literal: " + wireValue);
    }
    return java.util.Objects.requireNonNull(BY_WIRE_VALUE.get(wireValue));
  }

  /** Returns whether one published wire token is a valid stored-cell error literal. */
  public static boolean isSupportedWireValue(String wireValue) {
    if (wireValue == null) {
      throw new IllegalArgumentException("wireValue must not be null");
    }
    ExcelStoredCellErrorLiteral literal = BY_WIRE_VALUE.get(wireValue);
    return literal != null;
  }

  /** Returns the published stored error tokens in stable wire order. */
  public static List<String> orderedWireValues() {
    return Arrays.stream(values()).map(ExcelStoredCellErrorLiteral::wireValue).toList();
  }
}
