package dev.erst.gridgrind.protocol;

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

  /** One sheet-owned autofilter stored directly on the worksheet. */
  record SheetOwned(String range) implements AutofilterEntryReport {
    public SheetOwned {
      Objects.requireNonNull(range, "range must not be null");
    }
  }

  /** One table-owned autofilter stored on a table definition. */
  record TableOwned(String range, String tableName) implements AutofilterEntryReport {
    public TableOwned {
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(tableName, "tableName must not be null");
    }
  }
}
