package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.query.CellReadFacet;
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
    if (type == CellReadFacet.class) {
      return CELL_READ_FACET_DOCS;
    }
    return List.of();
  }

  static List<EnumValueDocEntry> orderedEnumDocs(
      Class<? extends Enum<?>> enumType, Map<String, String> docsByValue) {
    List<String> expectedValues =
        Arrays.stream(enumType.getEnumConstants()).map(Enum::name).toList();
    if (!docsByValue.keySet().equals(java.util.Set.copyOf(expectedValues))) {
      throw new IllegalStateException(
          "Enum value docs must cover every published token for " + enumType.getName());
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
