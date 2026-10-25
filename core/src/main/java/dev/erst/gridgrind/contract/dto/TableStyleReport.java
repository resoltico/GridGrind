package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Protocol-facing factual table-style metadata loaded from a workbook. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TableStyleReport.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = TableStyleReport.Named.class, name = "NAMED")
})
public sealed interface TableStyleReport permits TableStyleReport.None, TableStyleReport.Named {

  /** Table has no style metadata. */
  record None() implements TableStyleReport {}

  /** Table carries a named style plus explicit stripe and emphasis flags. */
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
