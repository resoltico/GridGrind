package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Repeating print-title column state captured from one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintTitleColumnsReport.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PrintTitleColumnsReport.Band.class, name = "BAND")
})
public sealed interface PrintTitleColumnsReport
    permits PrintTitleColumnsReport.None, PrintTitleColumnsReport.Band {
  /** Sheet has no repeating print-title columns. */
  record None() implements PrintTitleColumnsReport {}

  /** Sheet repeats the provided inclusive zero-based column band on every printed page. */
  record Band(int firstColumnIndex, int lastColumnIndex) implements PrintTitleColumnsReport {}
}
