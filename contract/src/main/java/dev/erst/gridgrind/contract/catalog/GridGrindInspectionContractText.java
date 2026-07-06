package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.stream.Collectors;

/** Inspection, analysis, and formula-surface wording shared by public contract surfaces. */
public final class GridGrindInspectionContractText {
  private static final List<String> WORKBOOK_ANALYSIS_FAMILIES =
      List.of(
          "formula health",
          "data-validation health",
          "conditional-formatting health",
          "autofilter health",
          "table health",
          "pivot-table health",
          "hyperlink health",
          "named-range health");

  public static final String FORMULA_SURFACE_READ_SUMMARY =
      "Return surface.totalFormulaCellCount plus per-sheet formula usage groups."
          + " surface.sheets[*] includes sheetName, formulaCellCount,"
          + " distinctFormulaCount, and grouped formulas with occurrenceCount"
          + " and addresses.";
  public static final String NAMED_RANGE_SURFACE_READ_SUMMARY =
      "Return surface.workbookScopedCount, sheetScopedCount,"
          + " rangeBackedCount, formulaBackedCount, and namedRanges."
          + " Each namedRanges entry reports name, scope, refersToFormula,"
          + " and backing kind.";
  public static final String FORMULA_HEALTH_READ_SUMMARY =
      "Return analysis.checkedFormulaCellCount, a severity summary,"
          + " and findings for formula errors, volatile usage,"
          + " or evaluation failures.";
  public static final String NAMED_RANGE_HEALTH_READ_SUMMARY =
      "Return analysis.checkedNamedRangeCount, a severity summary,"
          + " and named-range findings such as broken references,"
          + " unresolved targets, or scope shadowing.";
  public static final String WORKBOOK_FINDINGS_READ_SUMMARY =
      "Return analysis.summary plus one flat analysis.findings list after running"
          + " all analysis families (formula health, data-validation health,"
          + " conditional-formatting health, autofilter health, table health,"
          + " pivot-table health, hyperlink health, named-range health)"
          + " across the entire workbook and aggregate findings in a single response."
          + " This is the primary workbook-health check and pairs naturally with"
          + " persistence.type=NONE when no save is required.";

  private GridGrindInspectionContractText() {}

  /** Human-readable aggregate analysis-family list used by workbook-health discovery surfaces. */
  public static String workbookAnalysisFamilyPhrase() {
    return humanJoin(WORKBOOK_ANALYSIS_FAMILIES);
  }

  /** One stable description of request-authored formula boundaries. */
  public static String formulaAuthoringLimitSummary() {
    return "request-authored formulas are scalar only; array-formula braces such as"
        + " {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA, and authored LAMBDA/LET"
        + " currently surface as UNSUPPORTED_FORMULA_CONSTRUCT because Apache POI cannot"
        + " parse them on the write path.";
  }

  /** One stable description of loaded-formula evaluation boundaries. */
  public static String loadedFormulaSupportSummary() {
    return "formulas that Apache POI parses but cannot evaluate surface as"
        + " UNSUPPORTED_FORMULA.";
  }

  /** One stable catalog summary for `GET_SHEET_LAYOUT`. */
  public static String sheetLayoutReadSummary() {
    return "Return one sheet's layout object with pane, zoomPercent, presentation,"
        + " and per-row or per-column metadata."
        + " Row and column entries include hidden, outlineLevel, and collapsed"
        + " state where Excel persists it."
        + " Readback is factual and does not clamp malformed positive persisted"
        + " row heights, column widths, or default row height values.";
  }

  /** One stable catalog summary for `GET_FORMULA_SURFACE`. */
  public static String formulaSurfaceReadSummary() {
    return FORMULA_SURFACE_READ_SUMMARY;
  }

  /** One stable catalog summary for `GET_NAMED_RANGE_SURFACE`. */
  public static String namedRangeSurfaceReadSummary() {
    return NAMED_RANGE_SURFACE_READ_SUMMARY;
  }

  /** One stable catalog summary for `ANALYZE_FORMULA_HEALTH`. */
  public static String formulaHealthReadSummary() {
    return FORMULA_HEALTH_READ_SUMMARY;
  }

  /** One stable catalog summary for `ANALYZE_NAMED_RANGE_HEALTH`. */
  public static String namedRangeHealthReadSummary() {
    return NAMED_RANGE_HEALTH_READ_SUMMARY;
  }

  /** One stable catalog summary for `ANALYZE_WORKBOOK_FINDINGS`. */
  public static String workbookFindingsReadSummary() {
    return WORKBOOK_FINDINGS_READ_SUMMARY;
  }

  /** One stable discovery line for `ANALYZE_WORKBOOK_FINDINGS`. */
  public static String workbookFindingsDiscoverySummary() {
    return "ANALYZE_WORKBOOK_FINDINGS aggregates " + workbookAnalysisFamilyPhrase() + ".";
  }

  static String humanJoin(List<String> values) {
    List<String> parts = values.stream().filter(value -> !value.isBlank()).toList();
    if (parts.isEmpty()) {
      throw new IllegalArgumentException("values must not be empty");
    }
    if (parts.size() == 1) {
      return parts.getFirst();
    }
    if (parts.size() == 2) {
      return parts.getFirst() + " and " + parts.getLast();
    }
    return parts.subList(0, parts.size() - 1).stream().collect(Collectors.joining(", "))
        + ", and "
        + parts.getLast();
  }
}
