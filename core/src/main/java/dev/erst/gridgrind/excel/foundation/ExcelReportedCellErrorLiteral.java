package dev.erst.gridgrind.excel.foundation;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

/** GridGrind-owned wire vocabulary for reported cell error values. */
public enum ExcelReportedCellErrorLiteral {
  NULL("#NULL!"),
  DIV0("#DIV/0!"),
  VALUE("#VALUE!"),
  REF("#REF!"),
  NAME("#NAME?"),
  NUM("#NUM!"),
  NA("#N/A"),
  CIRCULAR_REF("#CIRCULAR_REF!"),
  FUNCTION_NOT_IMPLEMENTED("#FUNCTION_NOT_IMPLEMENTED!");

  private static final Map<String, ExcelReportedCellErrorLiteral> BY_WIRE_VALUE =
      Arrays.stream(values())
          .collect(
              java.util.stream.Collectors.toUnmodifiableMap(
                  ExcelReportedCellErrorLiteral::wireValue, literal -> literal));

  private final String wireValue;

  ExcelReportedCellErrorLiteral(String wireValue) {
    this.wireValue = wireValue;
  }

  /** Returns the published wire token for this error literal. */
  public String wireValue() {
    return wireValue;
  }

  /** Resolves one published wire token to its owned error literal. */
  public static ExcelReportedCellErrorLiteral fromWireValue(String wireValue) {
    if (wireValue == null) {
      throw new IllegalArgumentException("wireValue must not be null");
    }
    if (!isSupportedWireValue(wireValue)) {
      throw new IllegalArgumentException(
          "Unsupported reported Excel cell error literal: " + wireValue);
    }
    return java.util.Objects.requireNonNull(BY_WIRE_VALUE.get(wireValue));
  }

  /** Returns whether one published wire token is part of the owned error vocabulary. */
  public static boolean isSupportedWireValue(String wireValue) {
    if (wireValue == null) {
      throw new IllegalArgumentException("wireValue must not be null");
    }
    ExcelReportedCellErrorLiteral literal = BY_WIRE_VALUE.get(wireValue);
    return literal != null;
  }

  /** Returns the published error tokens in stable wire order. */
  public static List<String> orderedWireValues() {
    return Arrays.stream(values()).map(ExcelReportedCellErrorLiteral::wireValue).toList();
  }
}
