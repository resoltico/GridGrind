package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/** Shared note ids and payloads for the published protocol catalog. */
final class GridGrindProtocolCatalogNotes {
  static final String REQUEST_OWNED_PATH_RULE_ID = "requestOwnedPathRule";
  static final String FILE_HYPERLINK_RELATIVE_PATH_RULE_ID = "fileHyperlinkRelativePathRule";
  private static final List<String> REQUEST_OWNED_PATH_RULE_REF =
      List.of(REQUEST_OWNED_PATH_RULE_ID);
  private static final List<String> FILE_HYPERLINK_RELATIVE_PATH_RULE_REF =
      List.of(FILE_HYPERLINK_RELATIVE_PATH_RULE_ID);
  private static final List<CatalogNote> NOTES =
      List.of(
          new CatalogNote(
              REQUEST_OWNED_PATH_RULE_ID,
              GridGrindRequestSurfaceContractText.requestOwnedPathResolutionSummary()),
          new CatalogNote(
              FILE_HYPERLINK_RELATIVE_PATH_RULE_ID,
              "Relative FILE hyperlink paths resolve against the saved workbook's directory, not"
                  + " the request or execution root."));

  private GridGrindProtocolCatalogNotes() {}

  static List<CatalogNote> notes() {
    return NOTES;
  }

  static List<String> requestOwnedPathRuleRef() {
    return REQUEST_OWNED_PATH_RULE_REF;
  }

  static List<String> fileHyperlinkRelativePathRuleRef() {
    return FILE_HYPERLINK_RELATIVE_PATH_RULE_REF;
  }
}
