package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/** Shared note ids and payloads for the published protocol catalog. */
final class GridGrindProtocolCatalogNotes {
  static final String REQUEST_OWNED_PATH_RULE_ID = "requestOwnedPathRule";
  private static final List<String> REQUEST_OWNED_PATH_RULE_REF =
      List.of(REQUEST_OWNED_PATH_RULE_ID);
  private static final List<CatalogNote> NOTES =
      List.of(
          new CatalogNote(
              REQUEST_OWNED_PATH_RULE_ID,
              GridGrindRequestSurfaceContractText.requestOwnedPathResolutionSummary()));

  private GridGrindProtocolCatalogNotes() {}

  static List<CatalogNote> notes() {
    return NOTES;
  }

  static List<String> requestOwnedPathRuleRef() {
    return REQUEST_OWNED_PATH_RULE_REF;
  }
}
