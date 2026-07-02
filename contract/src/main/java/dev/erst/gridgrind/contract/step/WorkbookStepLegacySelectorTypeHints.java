package dev.erst.gridgrind.contract.step;

import java.util.Set;

/** Retired generic selector ids that need family-specific public guidance. */
final class WorkbookStepLegacySelectorTypeHints {
  private static final Set<String> RETIRED_GENERIC_TYPE_IDS =
      Set.of(
          "CURRENT",
          "ALL",
          "ALL_ON_SHEET",
          "ALL_ROWS",
          "ALL_USED_IN_SHEET",
          "ANY_OF",
          "BY_ADDRESS",
          "BY_ADDRESSES",
          "BY_COLUMN_NAME",
          "BY_INDEX",
          "BY_KEY_CELL",
          "BY_NAME",
          "BY_NAME_ON_SHEET",
          "BY_NAMES",
          "BY_QUALIFIED_ADDRESSES",
          "BY_RANGE",
          "BY_RANGES",
          "INSERTION",
          "RECTANGULAR_WINDOW",
          "SHEET_SCOPE",
          "SPAN",
          "WORKBOOK_SCOPE");

  private WorkbookStepLegacySelectorTypeHints() {}

  static String guidancePrefix(String authoredType) {
    return RETIRED_GENERIC_TYPE_IDS.contains(authoredType)
        ? "target selector ids are family-specific; "
        : "";
  }
}
