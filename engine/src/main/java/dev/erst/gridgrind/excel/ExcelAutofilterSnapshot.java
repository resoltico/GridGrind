package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Factual sheet- or table-owned autofilter metadata loaded from one sheet. */
public sealed interface ExcelAutofilterSnapshot
    permits ExcelAutofilterSnapshot.SheetOwned, ExcelAutofilterSnapshot.TableOwned {

  /** Raw A1-style filter range text as stored in workbook XML. */
  String range();

  /** One sheet-owned autofilter stored directly on the worksheet. */
  record SheetOwned(String range) implements ExcelAutofilterSnapshot {
    public SheetOwned {
      Objects.requireNonNull(range, "range must not be null");
    }
  }

  /** One table-owned autofilter stored on a table definition. */
  record TableOwned(String range, String tableName) implements ExcelAutofilterSnapshot {
    public TableOwned {
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(tableName, "tableName must not be null");
    }
  }
}
