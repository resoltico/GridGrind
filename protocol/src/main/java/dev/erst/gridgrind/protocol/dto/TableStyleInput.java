package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Supported table-style definition for GridGrind-authored tables. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TableStyleInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = TableStyleInput.Named.class, name = "NAMED")
})
public sealed interface TableStyleInput permits TableStyleInput.None, TableStyleInput.Named {

  /** Clear style metadata and leave the table unstyled. */
  record None() implements TableStyleInput {}

  /** Apply one named workbook table style with explicit stripe and emphasis flags. */
  record Named(
      String name,
      boolean showFirstColumn,
      boolean showLastColumn,
      boolean showRowStripes,
      boolean showColumnStripes)
      implements TableStyleInput {
    public Named {
      Objects.requireNonNull(name, "name must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
    }
  }
}
