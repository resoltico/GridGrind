package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Target placement for a copied sheet within workbook sheet order. */
@Union(discriminator = "type", catalog = true, displayOrder = 21)
public sealed interface SheetCopyPosition
    permits dev.erst.gridgrind.core.vocabulary.SheetCopyPosition.AppendAtEnd,
        dev.erst.gridgrind.core.vocabulary.SheetCopyPosition.AtIndex {

  /** Places the copied sheet after every existing sheet. */
  @Variant(id = "APPEND_AT_END", summary = "Place the copied sheet after every existing sheet.")
  record AppendAtEnd() implements SheetCopyPosition {}

  /** Places the copied sheet at the requested zero-based workbook position. */
  @Variant(
      id = "AT_INDEX",
      summary = "Place the copied sheet at the requested zero-based workbook position.")
  record AtIndex(int targetIndex) implements SheetCopyPosition {
    public AtIndex {
      if (targetIndex < 0) {
        throw new IllegalArgumentException("targetIndex must be >= 0");
      }
    }
  }
}
