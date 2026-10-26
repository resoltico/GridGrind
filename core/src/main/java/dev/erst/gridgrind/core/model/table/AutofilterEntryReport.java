package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import java.util.ArrayList;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual report for one sheet- or table-owned autofilter. */
@Union(discriminator = "type")
public sealed interface AutofilterEntryReport
    permits AutofilterEntryReport.SheetOwned, AutofilterEntryReport.TableOwned {

  /** Raw A1-style filter range text as stored in the workbook. */
  RangeLabel range();

  /** Persisted filter-column criteria carried by the autofilter. */
  java.util.List<AutofilterFilterColumnReport> filterColumns();

  /** Persisted sort-state metadata carried by the autofilter, when present. */
  Optional<AutofilterSortStateReport> sortState();

  /** One sheet-owned autofilter stored directly on the worksheet. */
  @Variant(id = "SHEET", summary = "One sheet-owned autofilter stored directly on the worksheet.")
  record SheetOwned(
      RangeLabel range,
      java.util.List<AutofilterFilterColumnReport> filterColumns,
      Optional<AutofilterSortStateReport> sortState)
      implements AutofilterEntryReport {
    /** Creates a sheet-owned autofilter report with no persisted criteria or sort state. */
    public SheetOwned(RangeLabel range) {
      this(range, java.util.List.of(), Optional.empty());
    }

    public SheetOwned {
      Objects.requireNonNull(range, "range must not be null");
      filterColumns = copyValues(filterColumns, "filterColumns");
      Objects.requireNonNull(sortState, "sortState must not be null");
    }
  }

  /** One table-owned autofilter stored on a table definition. */
  @Variant(id = "TABLE", summary = "One table-owned autofilter stored on a table definition.")
  record TableOwned(
      RangeLabel range,
      String tableName,
      java.util.List<AutofilterFilterColumnReport> filterColumns,
      Optional<AutofilterSortStateReport> sortState)
      implements AutofilterEntryReport {
    /** Creates a table-owned autofilter report with no persisted criteria or sort state. */
    public TableOwned(RangeLabel range, String tableName) {
      this(range, tableName, java.util.List.of(), Optional.empty());
    }

    public TableOwned {
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(tableName, "tableName must not be null");
      filterColumns = copyValues(filterColumns, "filterColumns");
      Objects.requireNonNull(sortState, "sortState must not be null");
    }
  }

  private static <T> java.util.List<T> copyValues(java.util.List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    java.util.List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return java.util.List.copyOf(copy);
  }
}
