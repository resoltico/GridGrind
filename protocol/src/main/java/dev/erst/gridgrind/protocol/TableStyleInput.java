package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import java.util.Objects;

/** Supported table-style definition for GridGrind-authored tables. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = TableStyleInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = TableStyleInput.Named.class, name = "NAMED")
})
public sealed interface TableStyleInput permits TableStyleInput.None, TableStyleInput.Named {

  /** Converts this protocol payload into the workbook-core table style. */
  ExcelTableStyle toExcelTableStyle();

  /** Clear style metadata and leave the table unstyled. */
  record None() implements TableStyleInput {
    @Override
    public ExcelTableStyle toExcelTableStyle() {
      return new ExcelTableStyle.None();
    }
  }

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

    @Override
    public ExcelTableStyle toExcelTableStyle() {
      return new ExcelTableStyle.Named(
          name, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes);
    }
  }
}
