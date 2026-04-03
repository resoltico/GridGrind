package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Protocol-facing scope of a named range, either workbook-wide or local to one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = NamedRangeScope.Workbook.class, name = "WORKBOOK"),
  @JsonSubTypes.Type(value = NamedRangeScope.Sheet.class, name = "SHEET")
})
public sealed interface NamedRangeScope permits NamedRangeScope.Workbook, NamedRangeScope.Sheet {

  /** Workbook-wide named-range scope. */
  record Workbook() implements NamedRangeScope {}

  /** Sheet-local named-range scope bound to one sheet name. */
  record Sheet(String sheetName) implements NamedRangeScope {
    public Sheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }
}
