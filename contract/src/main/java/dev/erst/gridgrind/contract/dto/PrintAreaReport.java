package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Print-area state captured from one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintAreaReport.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PrintAreaReport.Range.class, name = "RANGE")
})
public sealed interface PrintAreaReport permits PrintAreaReport.None, PrintAreaReport.Range {
  /** Sheet has no explicit print area. */
  record None() implements PrintAreaReport {}

  /** Sheet prints the provided rectangular A1-style range. */
  record Range(String range) implements PrintAreaReport {
    public Range {
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }
}
