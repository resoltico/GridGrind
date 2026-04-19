package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** One requested repeating print-title row band in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintTitleRowsInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PrintTitleRowsInput.Band.class, name = "BAND")
})
public sealed interface PrintTitleRowsInput
    permits PrintTitleRowsInput.None, PrintTitleRowsInput.Band {
  int MAX_ROW_INDEX = 1_048_575;

  /** Sheet has no repeating print-title rows. */
  record None() implements PrintTitleRowsInput {}

  /** Sheet repeats the provided inclusive zero-based row band on every printed page. */
  record Band(Integer firstRowIndex, Integer lastRowIndex) implements PrintTitleRowsInput {
    public Band {
      requireBand(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");
      if (lastRowIndex > MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            "lastRowIndex must not exceed " + MAX_ROW_INDEX + ": " + lastRowIndex);
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
