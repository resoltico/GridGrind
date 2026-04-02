package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One explicit sheet pane state captured or authored through the workbook model. */
public sealed interface ExcelSheetPane
    permits ExcelSheetPane.None, ExcelSheetPane.Frozen, ExcelSheetPane.Split {

  /** Sheet has no active pane split or freeze state. */
  record None() implements ExcelSheetPane {}

  /** Sheet is frozen at the provided split and visible-origin coordinates. */
  record Frozen(int splitColumn, int splitRow, int leftmostColumn, int topRow)
      implements ExcelSheetPane {
    public Frozen {
      requireNonNegative(splitColumn, "splitColumn");
      requireNonNegative(splitRow, "splitRow");
      requireNonNegative(leftmostColumn, "leftmostColumn");
      requireNonNegative(topRow, "topRow");
      requireNonZeroSplit(splitColumn, splitRow, "splitColumn", "splitRow");
      if (splitColumn == 0 && leftmostColumn != 0) {
        throw new IllegalArgumentException(
            "leftmostColumn must be 0 when splitColumn is 0: " + leftmostColumn);
      }
      if (splitRow == 0 && topRow != 0) {
        throw new IllegalArgumentException("topRow must be 0 when splitRow is 0: " + topRow);
      }
      if (splitColumn > 0 && leftmostColumn < splitColumn) {
        throw new IllegalArgumentException(
            "leftmostColumn must be greater than or equal to splitColumn");
      }
      if (splitRow > 0 && topRow < splitRow) {
        throw new IllegalArgumentException("topRow must be greater than or equal to splitRow");
      }
    }
  }

  /** Sheet uses split panes with explicit split offsets, visible origin, and active pane. */
  record Split(
      int xSplitPosition,
      int ySplitPosition,
      int leftmostColumn,
      int topRow,
      ExcelPaneRegion activePane)
      implements ExcelSheetPane {
    public Split {
      requireNonNegative(xSplitPosition, "xSplitPosition");
      requireNonNegative(ySplitPosition, "ySplitPosition");
      requireNonNegative(leftmostColumn, "leftmostColumn");
      requireNonNegative(topRow, "topRow");
      Objects.requireNonNull(activePane, "activePane must not be null");
      requireNonZeroSplit(xSplitPosition, ySplitPosition, "xSplitPosition", "ySplitPosition");
      if (xSplitPosition == 0 && leftmostColumn != 0) {
        throw new IllegalArgumentException(
            "leftmostColumn must be 0 when xSplitPosition is 0: " + leftmostColumn);
      }
      if (ySplitPosition == 0 && topRow != 0) {
        throw new IllegalArgumentException("topRow must be 0 when ySplitPosition is 0: " + topRow);
      }
    }
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static void requireNonZeroSplit(
      int primary, int secondary, String primaryFieldName, String secondaryFieldName) {
    if (primary == 0 && secondary == 0) {
      throw new IllegalArgumentException(
          primaryFieldName + " and " + secondaryFieldName + " must not both be 0");
    }
  }
}
