package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;

/** Selects one or more tables by workbook-global identity. */
@Union(discriminator = "type", catalog = true, displayOrder = 31)
public sealed interface TableSelector extends Selector
    permits dev.erst.gridgrind.core.plan.TableSelector.All,
        dev.erst.gridgrind.core.plan.TableSelector.ByName,
        dev.erst.gridgrind.core.plan.TableSelector.ByNames,
        dev.erst.gridgrind.core.plan.TableSelector.ByNameOnSheet {

  /** Selects every table in workbook order. */
  @Variant(id = "TABLE_ALL", summary = "Select every table in workbook order.")
  record All() implements TableSelector, ResolvedTarget {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one table by workbook-global name. */
  @Variant(id = "TABLE_BY_NAME", summary = "Select one workbook-global table by exact name.")
  record ByName(TableName name) implements TableSelector, ResolvedTarget {
    public ByName {
      java.util.Objects.requireNonNull(name, "name must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more tables by workbook-global names. */
  @Variant(
      id = "TABLE_BY_NAMES",
      summary = "Select one or more workbook-global tables by exact name.")
  record ByNames(List<TableName> names) implements TableSelector, ResolvedTarget {
    public ByNames {
      names = SelectorListValidation.copyDistinctTableNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one workbook-global table and asserts its expected owning sheet. */
  @Variant(
      id = "TABLE_BY_NAME_ON_SHEET",
      summary = "Select one workbook-global table by exact name and expected owning sheet.")
  record ByNameOnSheet(TableName name, SheetName sheetName)
      implements TableSelector, ResolvedTarget {
    public ByNameOnSheet {
      java.util.Objects.requireNonNull(name, "name must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
