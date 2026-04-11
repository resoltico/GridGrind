package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Factual sheet- or table-owned autofilter metadata loaded from one sheet. */
public sealed interface ExcelAutofilterSnapshot
    permits ExcelAutofilterSnapshot.SheetOwned, ExcelAutofilterSnapshot.TableOwned {

  /** Raw A1-style filter range text as stored in workbook XML. */
  String range();

  /** Persisted filter-column criteria carried by the autofilter. */
  java.util.List<ExcelAutofilterFilterColumnSnapshot> filterColumns();

  /** Persisted sort-state metadata carried by the autofilter, when present. */
  ExcelAutofilterSortStateSnapshot sortState();

  /** One sheet-owned autofilter stored directly on the worksheet. */
  record SheetOwned(
      String range,
      java.util.List<ExcelAutofilterFilterColumnSnapshot> filterColumns,
      ExcelAutofilterSortStateSnapshot sortState)
      implements ExcelAutofilterSnapshot {
    /** Creates a sheet-owned autofilter with no persisted criteria or sort state. */
    public SheetOwned(String range) {
      this(range, java.util.List.of(), null);
    }

    public SheetOwned {
      Objects.requireNonNull(range, "range must not be null");
      filterColumns =
          java.util.List.copyOf(
              Objects.requireNonNull(filterColumns, "filterColumns must not be null"));
      for (ExcelAutofilterFilterColumnSnapshot filterColumn : filterColumns) {
        Objects.requireNonNull(filterColumn, "filterColumns must not contain null values");
      }
    }
  }

  /** One table-owned autofilter stored on a table definition. */
  record TableOwned(
      String range,
      String tableName,
      java.util.List<ExcelAutofilterFilterColumnSnapshot> filterColumns,
      ExcelAutofilterSortStateSnapshot sortState)
      implements ExcelAutofilterSnapshot {
    /** Creates a table-owned autofilter with no persisted criteria or sort state. */
    public TableOwned(String range, String tableName) {
      this(range, tableName, java.util.List.of(), null);
    }

    public TableOwned {
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(tableName, "tableName must not be null");
      filterColumns =
          java.util.List.copyOf(
              Objects.requireNonNull(filterColumns, "filterColumns must not be null"));
      for (ExcelAutofilterFilterColumnSnapshot filterColumn : filterColumns) {
        Objects.requireNonNull(filterColumn, "filterColumns must not contain null values");
      }
    }
  }
}
