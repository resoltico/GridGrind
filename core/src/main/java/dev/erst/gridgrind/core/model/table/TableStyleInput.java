package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Supported table-style definition for GridGrind-authored tables. */
@Union(discriminator = "type", catalog = true, displayOrder = 70)
public sealed interface TableStyleInput
    permits dev.erst.gridgrind.core.model.table.TableStyleInput.None,
        dev.erst.gridgrind.core.model.table.TableStyleInput.Named {

  /** Clear style metadata and leave the table unstyled. */
  @Variant(id = "NONE", summary = "Clear table style metadata.")
  record None() implements TableStyleInput {}

  /** Apply one named workbook table style with explicit stripe and emphasis flags. */
  @Variant(
      id = "NAMED",
      summary = "Apply one named workbook table style with explicit stripe and emphasis flags.")
  record Named(
      String name,
      boolean showFirstColumn,
      boolean showLastColumn,
      boolean showRowStripes,
      boolean showColumnStripes)
      implements TableStyleInput {
    public Named {
      Objects.requireNonNull(name, "name must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
    }
  }
}
