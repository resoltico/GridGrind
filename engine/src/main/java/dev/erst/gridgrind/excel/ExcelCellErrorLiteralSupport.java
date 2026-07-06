package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelReportedCellErrorLiteral;
import dev.erst.gridgrind.excel.foundation.ExcelStoredCellErrorLiteral;
import java.util.Map;
import org.apache.poi.ss.usermodel.FormulaError;

/** Bridges GridGrind-owned stored and reported error literals to and from Apache POI. */
public final class ExcelCellErrorLiteralSupport {
  private static final Map<FormulaError, String> TO_REPORTED_WIRE_VALUE =
      Map.ofEntries(
          Map.entry(FormulaError.NULL, ExcelReportedCellErrorLiteral.NULL.wireValue()),
          Map.entry(FormulaError.DIV0, ExcelReportedCellErrorLiteral.DIV0.wireValue()),
          Map.entry(FormulaError.VALUE, ExcelReportedCellErrorLiteral.VALUE.wireValue()),
          Map.entry(FormulaError.REF, ExcelReportedCellErrorLiteral.REF.wireValue()),
          Map.entry(FormulaError.NAME, ExcelReportedCellErrorLiteral.NAME.wireValue()),
          Map.entry(FormulaError.NUM, ExcelReportedCellErrorLiteral.NUM.wireValue()),
          Map.entry(FormulaError.NA, ExcelReportedCellErrorLiteral.NA.wireValue()),
          Map.entry(
              FormulaError.CIRCULAR_REF, ExcelReportedCellErrorLiteral.CIRCULAR_REF.wireValue()),
          Map.entry(
              FormulaError.FUNCTION_NOT_IMPLEMENTED,
              ExcelReportedCellErrorLiteral.FUNCTION_NOT_IMPLEMENTED.wireValue()));
  private static final Map<ExcelStoredCellErrorLiteral, FormulaError> TO_POI_STORED_ERROR =
      Map.ofEntries(
          Map.entry(ExcelStoredCellErrorLiteral.NULL, FormulaError.NULL),
          Map.entry(ExcelStoredCellErrorLiteral.DIV0, FormulaError.DIV0),
          Map.entry(ExcelStoredCellErrorLiteral.VALUE, FormulaError.VALUE),
          Map.entry(ExcelStoredCellErrorLiteral.REF, FormulaError.REF),
          Map.entry(ExcelStoredCellErrorLiteral.NAME, FormulaError.NAME),
          Map.entry(ExcelStoredCellErrorLiteral.NUM, FormulaError.NUM),
          Map.entry(ExcelStoredCellErrorLiteral.NA, FormulaError.NA));

  private ExcelCellErrorLiteralSupport() {}

  /** Returns the published GridGrind reported token for one Apache POI cell error. */
  public static String toReportedWireValue(FormulaError formulaError) {
    java.util.Objects.requireNonNull(formulaError, "formulaError must not be null");
    String wireValue = TO_REPORTED_WIRE_VALUE.get(formulaError);
    if (wireValue != null) {
      return wireValue;
    }
    throw new IllegalArgumentException(formulaError + " is not a publishable cell error");
  }

  /** Returns the Apache POI stored cell error backing one authorable GridGrind wire token. */
  public static FormulaError toPoiStoredFormulaError(String wireValue) {
    FormulaError formulaError =
        TO_POI_STORED_ERROR.get(ExcelStoredCellErrorLiteral.fromWireValue(wireValue));
    return java.util.Objects.requireNonNull(formulaError, "formulaError must not be null");
  }
}
