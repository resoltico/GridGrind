package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Observed defined-name scope, preserving stored sheet labels without authoring validation. */
@Union(discriminator = "type", catalog = true, displayOrder = 36)
public sealed interface NamedRangeScopeReport {
  /** Workbook-wide observed scope. */
  @Variant(id = "WORKBOOK", summary = "Observed workbook scope.")
  record Workbook() implements NamedRangeScopeReport {}

  /** Sheet-local observed scope. */
  @Variant(id = "SHEET", summary = "Observed local sheet scope.")
  record Sheet(SheetLabel sheetName) implements NamedRangeScopeReport {
    public Sheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
    }
  }
}
