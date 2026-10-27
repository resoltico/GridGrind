package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Supported table-style definition for GridGrind-authored tables. */
@Union(discriminator = "type", catalog = true, displayOrder = 70)
public sealed interface TableStyleSpec
    permits dev.erst.gridgrind.core.model.table.TableStyleSpec.None,
        dev.erst.gridgrind.core.model.table.TableStyleSpec.Named {

  /** Clear style metadata and leave the table unstyled. */
  @Variant(id = "NONE", summary = "Clear table style metadata.")
  record None() implements TableStyleSpec {}

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
      implements TableStyleSpec {
    public Named {
      name = TextValidation.requireNonBlank(name, "name");
    }
  }
}
