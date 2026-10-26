package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Protocol-facing scope of a named range, either workbook-wide or local to one sheet. */
@Union(discriminator = "type", catalog = true, displayOrder = 35)
public sealed interface NamedRangeScope
    permits dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook,
        dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet {

  /** Workbook-wide named-range scope. */
  @Variant(id = "WORKBOOK", summary = "Target workbook scope.")
  record Workbook() implements NamedRangeScope {}

  /** Sheet-local named-range scope bound to one sheet name. */
  @Variant(id = "SHEET", summary = "Target one specific sheet scope.")
  record Sheet(SheetName sheetName) implements NamedRangeScope {
    public Sheet {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }
  }
}
