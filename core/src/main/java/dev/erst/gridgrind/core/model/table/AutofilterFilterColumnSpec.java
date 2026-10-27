package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import java.util.Objects;

/** One authored autofilter filter-column payload. */
@CatalogType(
    id = "AutofilterFilterColumnSpec",
    summary = "One authored autofilter filter-column payload with an explicit column criterion.",
    displayOrder = 56)
public record AutofilterFilterColumnSpec(
    long columnId,
    @ProtocolField(optional = true) @WireDefault("true") boolean showButton,
    AutofilterFilterCriterionSpec criterion) {
  /** Creates a filter-column payload with an explicit visible button. */
  public static AutofilterFilterColumnSpec visibleButton(
      long columnId, AutofilterFilterCriterionSpec criterion) {
    return new AutofilterFilterColumnSpec(columnId, true, criterion);
  }

  public AutofilterFilterColumnSpec {
    if (columnId < 0L) {
      throw new IllegalArgumentException("columnId must not be negative");
    }
    Objects.requireNonNull(criterion, "criterion must not be null");
  }
}
