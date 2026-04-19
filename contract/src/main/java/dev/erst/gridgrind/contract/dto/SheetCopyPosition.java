package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Target placement for a copied sheet within workbook sheet order. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = SheetCopyPosition.AppendAtEnd.class, name = "APPEND_AT_END"),
  @JsonSubTypes.Type(value = SheetCopyPosition.AtIndex.class, name = "AT_INDEX")
})
public sealed interface SheetCopyPosition
    permits SheetCopyPosition.AppendAtEnd, SheetCopyPosition.AtIndex {

  /** Places the copied sheet after every existing sheet. */
  record AppendAtEnd() implements SheetCopyPosition {}

  /** Places the copied sheet at the requested zero-based workbook position. */
  record AtIndex(Integer targetIndex) implements SheetCopyPosition {
    public AtIndex {
      java.util.Objects.requireNonNull(targetIndex, "targetIndex must not be null");
      if (targetIndex < 0) {
        throw new IllegalArgumentException("targetIndex must be >= 0");
      }
    }
  }
}
