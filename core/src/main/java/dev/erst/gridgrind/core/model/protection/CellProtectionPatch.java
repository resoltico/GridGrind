package dev.erst.gridgrind.core.model.protection;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing cell protection patch used by {@link CellStylePatch}. */
@CatalogType(
    id = "CellProtectionPatch",
    summary =
        "Cell protection patch; at least one field must be set. These flags matter when sheet protection is enabled.",
    displayOrder = 51)
public record CellProtectionPatch(Optional<Boolean> locked, Optional<Boolean> hiddenFormula) {
  public CellProtectionPatch {
    Objects.requireNonNull(locked, "locked must not be null");
    Objects.requireNonNull(hiddenFormula, "hiddenFormula must not be null");
    if (locked.isEmpty() && hiddenFormula.isEmpty()) {
      throw new IllegalArgumentException("protection must set at least one attribute");
    }
  }
}
