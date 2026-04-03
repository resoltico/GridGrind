package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Repeating print-title row state captured from one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintTitleRowsReport.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PrintTitleRowsReport.Band.class, name = "BAND")
})
public sealed interface PrintTitleRowsReport
    permits PrintTitleRowsReport.None, PrintTitleRowsReport.Band {
  /** Sheet has no repeating print-title rows. */
  record None() implements PrintTitleRowsReport {}

  /** Sheet repeats the provided inclusive zero-based row band on every printed page. */
  record Band(int firstRowIndex, int lastRowIndex) implements PrintTitleRowsReport {}
}
