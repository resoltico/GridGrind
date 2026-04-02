package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;

/** Target placement for a copied sheet within workbook sheet order. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = SheetCopyPosition.AppendAtEnd.class, name = "APPEND_AT_END"),
  @JsonSubTypes.Type(value = SheetCopyPosition.AtIndex.class, name = "AT_INDEX")
})
public sealed interface SheetCopyPosition
    permits SheetCopyPosition.AppendAtEnd, SheetCopyPosition.AtIndex {

  /** Converts this protocol position into the workbook-core copy-position model. */
  ExcelSheetCopyPosition toExcelSheetCopyPosition();

  /** Places the copied sheet after every existing sheet. */
  record AppendAtEnd() implements SheetCopyPosition {
    @Override
    public ExcelSheetCopyPosition toExcelSheetCopyPosition() {
      return new ExcelSheetCopyPosition.AppendAtEnd();
    }
  }

  /** Places the copied sheet at the requested zero-based workbook position. */
  record AtIndex(Integer targetIndex) implements SheetCopyPosition {
    public AtIndex {
      WorkbookOperation.Validation.requireNonNegative(
          java.util.Objects.requireNonNull(targetIndex, "targetIndex must not be null"),
          "targetIndex");
    }

    @Override
    public ExcelSheetCopyPosition toExcelSheetCopyPosition() {
      return new ExcelSheetCopyPosition.AtIndex(targetIndex);
    }
  }
}
