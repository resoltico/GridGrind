package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.query.CellReadFacet;
import dev.erst.gridgrind.excel.foundation.ExcelReportedCellErrorLiteral;
import dev.erst.gridgrind.excel.foundation.ExcelStoredCellErrorLiteral;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Canonical machine-readable enum value summaries for catalog fields that need them. */
final class CatalogEnumValueDocumentationSupport {
  private static final List<EnumValueDocEntry> CELL_READ_FACET_DOCS =
      orderedEnumDocs(
          CellReadFacet.class,
          Map.of(
              "VALUE",
                  "Project the factual cell value: textValue, numberValue, booleanValue,"
                      + " errorValue, or formula evaluation.",
              "STYLE", "Project the style report for each returned cell.",
              "FORMAT", "Project displayValue using Excel's formatted display text.",
              "HYPERLINK", "Project hyperlink metadata when the cell carries a hyperlink.",
              "COMMENT", "Project comment metadata when the cell carries a comment.",
              "FORMULA", "Project authored formula text for formula cells.",
              "RICH_TEXT_RUNS",
                  "Project rich-text runs for text cells and text-valued formula evaluations.",
              "TEMPORAL",
                  "Project derived date, time, or date-time semantics for date-like numeric"
                      + " values."));
  private static final List<EnumValueDocEntry> STORED_CELL_ERROR_LITERAL_DOCS =
      orderedValueDocs(
          ExcelStoredCellErrorLiteral.orderedWireValues(),
          Map.of(
              "#NULL!", "Excel null-intersection error.",
              "#DIV/0!", "Excel division-by-zero error.",
              "#VALUE!", "Excel wrong-type or wrong-shape value error.",
              "#REF!", "Excel invalid-reference error.",
              "#NAME?", "Excel unknown-name error.",
              "#NUM!", "Excel numeric-domain error.",
              "#N/A", "Excel not-available error."),
          "Stored Excel cell error literal");
  private static final List<EnumValueDocEntry> REPORTED_CELL_ERROR_LITERAL_DOCS =
      orderedValueDocs(
          ExcelReportedCellErrorLiteral.orderedWireValues(),
          Map.of(
              "#NULL!", "Excel null-intersection error.",
              "#DIV/0!", "Excel division-by-zero error.",
              "#VALUE!", "Excel wrong-type or wrong-shape value error.",
              "#REF!", "Excel invalid-reference error.",
              "#NAME?", "Excel unknown-name error.",
              "#NUM!", "Excel numeric-domain error.",
              "#N/A", "Excel not-available error.",
              "#CIRCULAR_REF!", "GridGrind-owned circular-reference evaluation error.",
              "#FUNCTION_NOT_IMPLEMENTED!",
                  "GridGrind-owned unsupported-function evaluation error."),
          "Reported Excel cell error literal");

  private CatalogEnumValueDocumentationSupport() {}

  static List<EnumValueDocEntry> enumValueDocs(Type type) {
    Objects.requireNonNull(type, "type must not be null");
    if (type instanceof ParameterizedType parameterizedType
        && parameterizedType.getRawType() == java.util.Optional.class) {
      return enumValueDocs(
          CatalogFieldMetadataSupport.singleTypeArgument(parameterizedType, "Optional"));
    }
    if (type instanceof ParameterizedType parameterizedType
        && parameterizedType.getRawType() == java.util.List.class) {
      return enumValueDocs(
          CatalogFieldMetadataSupport.singleTypeArgument(parameterizedType, "List"));
    }
    if (type == ExcelStoredCellErrorLiteral.class) {
      return STORED_CELL_ERROR_LITERAL_DOCS;
    }
    if (type == ExcelReportedCellErrorLiteral.class) {
      return REPORTED_CELL_ERROR_LITERAL_DOCS;
    }
    if (type == CellReadFacet.class) {
      return CELL_READ_FACET_DOCS;
    }
    return List.of();
  }

  static List<EnumValueDocEntry> orderedEnumDocs(
      Class<? extends Enum<?>> enumType, Map<String, String> docsByValue) {
    List<String> expectedValues =
        Arrays.stream(enumType.getEnumConstants()).map(Enum::name).toList();
    return orderedValueDocs(expectedValues, docsByValue, enumType.getName());
  }

  static List<EnumValueDocEntry> storedCellErrorLiteralDocs() {
    return STORED_CELL_ERROR_LITERAL_DOCS;
  }

  static List<EnumValueDocEntry> reportedCellErrorLiteralDocs() {
    return REPORTED_CELL_ERROR_LITERAL_DOCS;
  }

  private static List<EnumValueDocEntry> orderedValueDocs(
      List<String> expectedValues, Map<String, String> docsByValue, String valueSetName) {
    if (!docsByValue.keySet().equals(java.util.Set.copyOf(expectedValues))) {
      throw new IllegalStateException(
          "Enum value docs must cover every published token for " + valueSetName);
    }
    return expectedValues.stream()
        .map(
            value ->
                new EnumValueDocEntry(
                    value,
                    Objects.requireNonNull(
                        docsByValue.get(value), "Missing enum value doc for " + value)))
        .toList();
  }
}
