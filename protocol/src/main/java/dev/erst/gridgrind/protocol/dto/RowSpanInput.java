package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelRowSpan;
import java.util.Objects;

/** Inclusive zero-based row band used by structural row operations. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({@JsonSubTypes.Type(value = RowSpanInput.Band.class, name = "BAND")})
public sealed interface RowSpanInput permits RowSpanInput.Band {
  /** Inclusive zero-based row band. */
  record Band(Integer firstRowIndex, Integer lastRowIndex) implements RowSpanInput {
    public Band {
      Objects.requireNonNull(firstRowIndex, "firstRowIndex must not be null");
      Objects.requireNonNull(lastRowIndex, "lastRowIndex must not be null");
      if (firstRowIndex < 0) {
        throw new IllegalArgumentException("firstRowIndex must not be negative");
      }
      if (lastRowIndex < 0) {
        throw new IllegalArgumentException("lastRowIndex must not be negative");
      }
      if (lastRowIndex < firstRowIndex) {
        throw new IllegalArgumentException("lastRowIndex must not be less than firstRowIndex");
      }
      if (firstRowIndex > ExcelRowSpan.MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            "firstRowIndex must not exceed " + ExcelRowSpan.MAX_ROW_INDEX + " (Excel row limit)");
      }
      if (lastRowIndex > ExcelRowSpan.MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            "lastRowIndex must not exceed " + ExcelRowSpan.MAX_ROW_INDEX + " (Excel row limit)");
      }
    }
  }
}
