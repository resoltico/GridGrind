package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;

/** Selects one or more pivot tables by workbook-global identity. */
@Union(discriminator = "type", catalog = true, displayOrder = 39)
public sealed interface PivotTableSelector extends Selector
    permits dev.erst.gridgrind.core.plan.PivotTableSelector.All,
        dev.erst.gridgrind.core.plan.PivotTableSelector.ByName,
        dev.erst.gridgrind.core.plan.PivotTableSelector.ByNames,
        dev.erst.gridgrind.core.plan.PivotTableSelector.ByNameOnSheet {

  /** Selects every pivot table in workbook order. */
  @Variant(id = "PIVOT_TABLE_ALL", summary = "Select every pivot table in workbook order.")
  record All() implements PivotTableSelector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ANY_NUMBER;
    }
  }

  /** Selects one pivot table by workbook-global pivot name. */
  @Variant(
      id = "PIVOT_TABLE_BY_NAME",
      summary = "Select one workbook-global pivot table by exact name.")
  record ByName(String name) implements PivotTableSelector {
    public ByName {
      name = SelectorValueValidation.requirePivotTableName(name, "name");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  /** Selects one or more pivot tables by workbook-global pivot names. */
  @Variant(
      id = "PIVOT_TABLE_BY_NAMES",
      summary = "Select one or more workbook-global pivot tables by exact name.")
  record ByNames(List<String> names) implements PivotTableSelector {
    public ByNames {
      names = SelectorListValidation.copyDistinctPivotTableNames(names, "names");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.ONE_OR_MORE;
    }
  }

  /** Selects one pivot table by workbook-global name and asserts its owning sheet. */
  @Variant(
      id = "PIVOT_TABLE_BY_NAME_ON_SHEET",
      summary = "Select one workbook-global pivot table by exact name and expected owning sheet.")
  record ByNameOnSheet(String name, SheetName sheetName) implements PivotTableSelector {
    public ByNameOnSheet {
      name = SelectorValueValidation.requirePivotTableName(name, "name");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    }

    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }
}
