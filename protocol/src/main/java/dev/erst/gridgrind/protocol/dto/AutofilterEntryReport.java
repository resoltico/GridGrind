package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Protocol-facing factual report for one sheet- or table-owned autofilter. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = AutofilterEntryReport.SheetOwned.class, name = "SHEET"),
  @JsonSubTypes.Type(value = AutofilterEntryReport.TableOwned.class, name = "TABLE")
})
public sealed interface AutofilterEntryReport
    permits AutofilterEntryReport.SheetOwned, AutofilterEntryReport.TableOwned {

  /** Raw A1-style filter range text as stored in the workbook. */
  String range();

  /** Persisted filter-column criteria carried by the autofilter. */
  java.util.List<AutofilterFilterColumnReport> filterColumns();

  /** Persisted sort-state metadata carried by the autofilter, when present. */
  AutofilterSortStateReport sortState();

  /** One sheet-owned autofilter stored directly on the worksheet. */
  record SheetOwned(
      String range,
      java.util.List<AutofilterFilterColumnReport> filterColumns,
      AutofilterSortStateReport sortState)
      implements AutofilterEntryReport {
    /** Creates a sheet-owned autofilter report with no persisted criteria or sort state. */
    public SheetOwned(String range) {
      this(range, java.util.List.of(), null);
    }

    public SheetOwned {
      Objects.requireNonNull(range, "range must not be null");
      filterColumns =
          java.util.List.copyOf(
              Objects.requireNonNull(filterColumns, "filterColumns must not be null"));
      for (AutofilterFilterColumnReport filterColumn : filterColumns) {
        Objects.requireNonNull(filterColumn, "filterColumns must not contain null values");
      }
    }
  }

  /** One table-owned autofilter stored on a table definition. */
  record TableOwned(
      String range,
      String tableName,
      java.util.List<AutofilterFilterColumnReport> filterColumns,
      AutofilterSortStateReport sortState)
      implements AutofilterEntryReport {
    /** Creates a table-owned autofilter report with no persisted criteria or sort state. */
    public TableOwned(String range, String tableName) {
      this(range, tableName, java.util.List.of(), null);
    }

    public TableOwned {
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(tableName, "tableName must not be null");
      filterColumns =
          java.util.List.copyOf(
              Objects.requireNonNull(filterColumns, "filterColumns must not be null"));
      for (AutofilterFilterColumnReport filterColumn : filterColumns) {
        Objects.requireNonNull(filterColumn, "filterColumns must not contain null values");
      }
    }
  }
}
