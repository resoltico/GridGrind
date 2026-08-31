package dev.erst.gridgrind.contract.catalog;

import java.util.Map;
import java.util.Set;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.JsonNodeFactory;
import tools.jackson.databind.node.ObjectNode;

/**
 * Owns preference tables and placeholder defaults used while synthesizing catalog step templates.
 */
final class CatalogStepTemplateDefaults {
  private static final Set<String> TYPED_CELL_INPUT_TEMPLATE_GROUPS =
      Set.of("cellRowInputTypes", "cellGridInputTypes");
  private static final Set<String> PRIMARY_SELECTOR_TYPE_IDS =
      Set.of(
          "SHEET_BY_NAME",
          "CELL_BY_ADDRESS",
          "RANGE_BY_RANGE",
          "ROW_BAND_BY_INDEX",
          "COLUMN_BAND_BY_INDEX",
          "DRAWING_OBJECT_BY_NAME",
          "CHART_BY_NAME",
          "TABLE_BY_NAME",
          "PIVOT_TABLE_BY_NAME",
          "NAMED_RANGE_BY_NAME",
          "NAMED_RANGE_WORKBOOK_SCOPE",
          "TABLE_ROW_BY_KEY",
          "TABLE_CELL_BY_KEY");
  private static final Set<String> MULTI_SELECTOR_TYPE_IDS =
      Set.of(
          "SHEET_BY_NAMES",
          "CELL_BY_ADDRESSES",
          "RANGE_BY_RANGES",
          "ROW_BAND_BY_INDEXES",
          "COLUMN_BAND_BY_INDEXES");
  private static final Set<String> PRIMARY_ENTRY_TYPE_IDS =
      Set.of(
          "INLINE",
          "INLINE_TEXT",
          "TEXT",
          "WORKBOOK_CURRENT",
          "SHEET_BY_NAME",
          "CELL_BY_ADDRESS",
          "RANGE_BY_RANGE");
  private static final Set<String> SECONDARY_ENTRY_TYPE_IDS =
      Set.of("BOOLEAN", "NUMBER", "EXACT", "CURRENT", "NONE");
  private static final Set<String> TERTIARY_ENTRY_TYPE_IDS =
      Set.of("INLINE_BASE64", "CELL_BY_ADDRESSES", "SHEET_BY_NAMES", "RANGE_BY_RANGES");
  private static final Map<String, String> STRING_PLACEHOLDERS =
      Map.ofEntries(
          Map.entry("sheetName", "Sheet1"),
          Map.entry("name", "Sheet1"),
          Map.entry("sourceSheetName", "Sheet1"),
          Map.entry("targetSheetName", "Sheet1"),
          Map.entry("address", "A1"),
          Map.entry("topLeftAddress", "A1"),
          Map.entry("range", "A1:B2"),
          Map.entry("formula", "1"),
          Map.entry("formula1", "1"),
          Map.entry("formula2", "2"),
          Map.entry("path", "sample-path"),
          Map.entry("pkcs12Path", "certificate.p12"),
          Map.entry("base64Data", "AA=="),
          Map.entry("rgb", "#336699"),
          Map.entry("text", "Sample text"),
          Map.entry("title", "Sample title"),
          Map.entry("label", "Sample label"),
          Map.entry("description", "Sample description"),
          Map.entry("displayName", "Sample name"),
          Map.entry("fileName", "sample.bin"),
          Map.entry("planId", "sample-plan"),
          Map.entry("stepId", "sample-step"),
          Map.entry("email", "signer@example.com"),
          Map.entry("suggestedSignerEmail", "signer@example.com"),
          Map.entry("relationshipId", "rId1"));
  private static final Map<String, Integer> NUMBER_PLACEHOLDERS =
      Map.ofEntries(
          Map.entry("zoomPercent", 100),
          Map.entry("rowIndex", 0),
          Map.entry("columnIndex", 0),
          Map.entry("firstRowIndex", 0),
          Map.entry("lastRowIndex", 0),
          Map.entry("firstColumnIndex", 0),
          Map.entry("lastColumnIndex", 0),
          Map.entry("dx", 0),
          Map.entry("dy", 0),
          Map.entry("twips", 20));
  private static final Set<String> TRUE_BOOLEAN_PLACEHOLDERS =
      Set.of("visible", "locked", "hidden");
  private static final Map<String, TypeDefaultApplier> TYPE_SPECIFIC_DEFAULTS =
      buildTypeSpecificDefaults();
  private static final JsonNodeFactory JSON = JsonNodeFactory.instance;

  private CatalogStepTemplateDefaults() {}

  static int selectorPreference(String typeId) {
    if ("WORKBOOK_CURRENT".equals(typeId)) {
      return 0;
    }
    if (PRIMARY_SELECTOR_TYPE_IDS.contains(typeId)) {
      return 1;
    }
    return MULTI_SELECTOR_TYPE_IDS.contains(typeId) ? 2 : 3;
  }

  static int entryPreference(String typeId) {
    if (PRIMARY_ENTRY_TYPE_IDS.contains(typeId)) {
      return 0;
    }
    if (SECONDARY_ENTRY_TYPE_IDS.contains(typeId)) {
      return 1;
    }
    return TERTIARY_ENTRY_TYPE_IDS.contains(typeId) ? 2 : 3;
  }

  static int entryPreference(String groupName, String typeId) {
    if (TYPED_CELL_INPUT_TEMPLATE_GROUPS.contains(groupName) && "TYPED".equals(typeId)) {
      return -1;
    }
    return entryPreference(typeId);
  }

  static String stringPlaceholder(String fieldName) {
    return STRING_PLACEHOLDERS.getOrDefault(fieldName, "sample-" + fieldName.replace('_', '-'));
  }

  static int numberPlaceholder(String fieldName) {
    return NUMBER_PLACEHOLDERS.getOrDefault(fieldName, 1);
  }

  static boolean booleanPlaceholder(String fieldName) {
    return TRUE_BOOLEAN_PLACEHOLDERS.contains(fieldName);
  }

  static void applyTypeSpecificDefaults(
      Catalog catalog,
      String typeId,
      ObjectNode object,
      Set<String> recursionGuard,
      java.util.List<String> notes) {
    TypeDefaultApplier defaultApplier = TYPE_SPECIFIC_DEFAULTS.get(typeId);
    if (defaultApplier != null) {
      defaultApplier.apply(catalog, object, recursionGuard, notes);
    }
  }

  /** Applies one type-id-specific placeholder enrichment rule. */
  @FunctionalInterface
  interface TypeDefaultApplier {
    /** Mutates the placeholder object with the rule's extra defaults. */
    void apply(
        Catalog catalog,
        ObjectNode object,
        Set<String> recursionGuard,
        java.util.List<String> notes);
  }

  static Map<String, TypeDefaultApplier> buildTypeSpecificDefaults() {
    return Map.ofEntries(
        Map.entry(
            "SET_SHEET_ZOOM",
            (catalog, object, recursionGuard, notes) -> object.put("zoomPercent", 100)),
        Map.entry(
            "ChartSeriesInput",
            (catalog, object, recursionGuard, notes) ->
                object.set(
                    "title",
                    CatalogStepTemplateSupport.nestedGroupTemplate(
                        catalog, "chartTitleInputTypes", recursionGuard, notes))),
        Map.entry(
            "CellStylePatchInput",
            (catalog, object, recursionGuard, notes) -> object.put("numberFormat", "0.00")),
        Map.entry(
            "CustomXmlMappingLocator",
            (catalog, object, recursionGuard, notes) -> object.put("name", "Mapping1")),
        Map.entry(
            "DifferentialStyleInput",
            (catalog, object, recursionGuard, notes) -> object.put("bold", true)),
        Map.entry(
            "FORMULA_RULE",
            (catalog, object, recursionGuard, notes) -> object.put("stopIfTrue", true)),
        Map.entry(
            "CELL_VALUE_RULE",
            (catalog, object, recursionGuard, notes) -> object.put("stopIfTrue", true)),
        Map.entry(
            "TOP10_RULE",
            (catalog, object, recursionGuard, notes) -> object.put("stopIfTrue", true)),
        Map.entry(
            "FontHeightReport",
            (catalog, object, recursionGuard, notes) -> {
              object.put("twips", 20);
              object.put("points", 1);
            }),
        Map.entry(
            "SignatureLineInput",
            (catalog, object, recursionGuard, notes) -> object.put("caption", "Sign here")),
        Map.entry(
            "URL",
            (catalog, object, recursionGuard, notes) ->
                object.put("target", "https://example.com")),
        Map.entry(
            "EXPECT_ANALYSIS_MAX_SEVERITY", CatalogStepTemplateDefaults::applyAnalysisQueryDefault),
        Map.entry(
            "EXPECT_ANALYSIS_FINDING_PRESENT",
            CatalogStepTemplateDefaults::applyAnalysisQueryDefault),
        Map.entry(
            "EXPECT_ANALYSIS_FINDING_ABSENT",
            CatalogStepTemplateDefaults::applyAnalysisQueryDefault),
        Map.entry("ALL_OF", CatalogStepTemplateDefaults::applyAssertionCollectionDefault),
        Map.entry("ANY_OF", CatalogStepTemplateDefaults::applyAssertionCollectionDefault),
        Map.entry(
            "NOT",
            (catalog, object, recursionGuard, notes) ->
                object.set(
                    "assertion",
                    CatalogStepTemplateSupport.typeTemplateById(
                        catalog, "EXPECT_ANALYSIS_MAX_SEVERITY", recursionGuard, notes))));
  }

  private static void applyAnalysisQueryDefault(
      Catalog catalog,
      ObjectNode object,
      Set<String> recursionGuard,
      java.util.List<String> notes) {
    object.set(
        "query",
        CatalogStepTemplateSupport.typeTemplateById(
            catalog, "ANALYZE_WORKBOOK_FINDINGS", recursionGuard, notes));
  }

  private static void applyAssertionCollectionDefault(
      Catalog catalog,
      ObjectNode object,
      Set<String> recursionGuard,
      java.util.List<String> notes) {
    ArrayNode assertions = JSON.arrayNode();
    assertions.add(
        CatalogStepTemplateSupport.typeTemplateById(
            catalog, "EXPECT_ANALYSIS_MAX_SEVERITY", recursionGuard, notes));
    object.set("assertions", assertions);
  }
}
