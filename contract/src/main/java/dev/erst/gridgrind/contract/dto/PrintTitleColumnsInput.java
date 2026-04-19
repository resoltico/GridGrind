package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** One requested repeating print-title column band in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintTitleColumnsInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PrintTitleColumnsInput.Band.class, name = "BAND")
})
public sealed interface PrintTitleColumnsInput
    permits PrintTitleColumnsInput.None, PrintTitleColumnsInput.Band {
  int MAX_COLUMN_INDEX = 16_383;

  /** Sheet has no repeating print-title columns. */
  record None() implements PrintTitleColumnsInput {}

  /** Sheet repeats the provided inclusive zero-based column band on every printed page. */
  record Band(Integer firstColumnIndex, Integer lastColumnIndex) implements PrintTitleColumnsInput {
    public Band {
      requireBand(firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");
      if (lastColumnIndex > MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            "lastColumnIndex must not exceed " + MAX_COLUMN_INDEX + ": " + lastColumnIndex);
      }
    }
  }

  private static void requireBand(
      Integer firstIndex, Integer lastIndex, String firstFieldName, String lastFieldName) {
    if (firstIndex == null) {
      throw new IllegalArgumentException(firstFieldName + " must not be null");
    }
    if (lastIndex == null) {
      throw new IllegalArgumentException(lastFieldName + " must not be null");
    }
    if (firstIndex < 0) {
      throw new IllegalArgumentException(firstFieldName + " must not be negative");
    }
    if (lastIndex < 0) {
      throw new IllegalArgumentException(lastFieldName + " must not be negative");
    }
    if (lastIndex < firstIndex) {
      throw new IllegalArgumentException(
          lastFieldName + " must not be less than " + firstFieldName);
    }
  }
}
