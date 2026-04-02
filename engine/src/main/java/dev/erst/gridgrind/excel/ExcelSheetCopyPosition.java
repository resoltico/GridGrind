package dev.erst.gridgrind.excel;

/** Target placement for a copied sheet within workbook sheet order. */
public sealed interface ExcelSheetCopyPosition
    permits ExcelSheetCopyPosition.AppendAtEnd, ExcelSheetCopyPosition.AtIndex {

  /** Places the copied sheet after every existing sheet. */
  record AppendAtEnd() implements ExcelSheetCopyPosition {}

  /** Places the copied sheet at the requested zero-based workbook position. */
  record AtIndex(int targetIndex) implements ExcelSheetCopyPosition {
    public AtIndex {
      if (targetIndex < 0) {
        throw new IllegalArgumentException("targetIndex must not be negative");
      }
    }
  }
}
