package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelColumnSpan;
import java.util.Objects;

/** Inclusive zero-based column band used by structural column operations. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({@JsonSubTypes.Type(value = ColumnSpanInput.Band.class, name = "BAND")})
public sealed interface ColumnSpanInput permits ColumnSpanInput.Band {
  /** Inclusive zero-based column band. */
  record Band(Integer firstColumnIndex, Integer lastColumnIndex) implements ColumnSpanInput {
    public Band {
      Objects.requireNonNull(firstColumnIndex, "firstColumnIndex must not be null");
      Objects.requireNonNull(lastColumnIndex, "lastColumnIndex must not be null");
      if (firstColumnIndex < 0) {
        throw new IllegalArgumentException("firstColumnIndex must not be negative");
      }
      if (lastColumnIndex < 0) {
        throw new IllegalArgumentException("lastColumnIndex must not be negative");
      }
      if (lastColumnIndex < firstColumnIndex) {
        throw new IllegalArgumentException(
            "lastColumnIndex must not be less than firstColumnIndex");
      }
      if (firstColumnIndex > ExcelColumnSpan.MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            "firstColumnIndex must not exceed "
                + ExcelColumnSpan.MAX_COLUMN_INDEX
                + " (Excel column limit)");
      }
      if (lastColumnIndex > ExcelColumnSpan.MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            "lastColumnIndex must not exceed "
                + ExcelColumnSpan.MAX_COLUMN_INDEX
                + " (Excel column limit)");
      }
    }
  }
}
