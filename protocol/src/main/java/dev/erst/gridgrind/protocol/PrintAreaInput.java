package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** One requested print-area state in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintAreaInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PrintAreaInput.Range.class, name = "RANGE")
})
public sealed interface PrintAreaInput permits PrintAreaInput.None, PrintAreaInput.Range {
  /** Sheet has no explicit print area. */
  record None() implements PrintAreaInput {}

  /** Sheet prints the provided rectangular A1-style range. */
  record Range(String range) implements PrintAreaInput {
    public Range {
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }
}
