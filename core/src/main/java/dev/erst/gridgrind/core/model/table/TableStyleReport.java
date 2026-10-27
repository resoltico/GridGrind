package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Protocol-facing factual table-style metadata loaded from a workbook. */
@Union(discriminator = "type", catalog = true, displayOrder = 8)
public sealed interface TableStyleReport
    permits dev.erst.gridgrind.core.model.table.TableStyleReport.None,
        dev.erst.gridgrind.core.model.table.TableStyleReport.Named {

  /** Table has no style metadata. */
  @Variant(id = "NONE", summary = "Expect the table to carry no persisted style.")
  record None() implements TableStyleReport {}

  /** Table carries a named style plus explicit stripe and emphasis flags. */
  @Variant(
      id = "NAMED",
      summary = "Expect the table to carry one named style plus stripe/emphasis flags.")
  record Named(
      String name,
      boolean showFirstColumn,
      boolean showLastColumn,
      boolean showRowStripes,
      boolean showColumnStripes)
      implements TableStyleReport {
    public Named {
      Objects.requireNonNull(name, "name must not be null");
    }
  }
}
