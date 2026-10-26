package dev.erst.gridgrind.core.model.protection;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Captures whether a sheet is protected and, if so, with which supported lock flags. */
@Union(discriminator = "kind", catalog = true, displayOrder = 7)
public sealed interface SheetProtectionReport
    permits dev.erst.gridgrind.core.model.protection.SheetProtectionReport.Unprotected,
        dev.erst.gridgrind.core.model.protection.SheetProtectionReport.Protected {
  /** Sheet protection is disabled. */
  @Variant(id = "UNPROTECTED", summary = "Expect the sheet to have no protection.")
  record Unprotected() implements SheetProtectionReport {}

  /** Sheet protection is enabled with the reported supported lock flags. */
  @Variant(
      id = "PROTECTED",
      summary = "Expect the sheet to be protected with explicit lock settings.")
  record Protected(SheetProtectionSettings settings) implements SheetProtectionReport {
    public Protected {
      Objects.requireNonNull(settings, "settings must not be null");
    }
  }
}
