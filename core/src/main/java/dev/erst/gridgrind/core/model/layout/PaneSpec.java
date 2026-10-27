package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import java.util.Objects;

/** One explicit sheet pane state requested through the wire protocol. */
@Union(discriminator = "type", catalog = true, displayOrder = 20)
public sealed interface PaneSpec
    permits dev.erst.gridgrind.core.model.layout.PaneSpec.None,
        dev.erst.gridgrind.core.model.layout.PaneSpec.Frozen,
        dev.erst.gridgrind.core.model.layout.PaneSpec.Split {
  /** Sheet has no active pane split or freeze state. */
  @Variant(id = "NONE", summary = "Sheet has no active pane split or freeze.")
  record None() implements PaneSpec {}

  /** Sheet is frozen at the provided split and visible-origin coordinates. */
  @Variant(
      id = "FROZEN",
      summary = "Freeze the sheet at the provided split and visible-origin coordinates.")
  record Frozen(int splitColumn, int splitRow, int leftmostColumn, int topRow) implements PaneSpec {
    public Frozen {
      Validation.requireNonNegative(splitColumn, "splitColumn");
      Validation.requireNonNegative(splitRow, "splitRow");
      Validation.requireNonNegative(leftmostColumn, "leftmostColumn");
      Validation.requireNonNegative(topRow, "topRow");
      Validation.requireNonZeroPair(splitColumn, splitRow, "splitColumn", "splitRow");
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
  @Variant(
      id = "SPLIT",
      summary = "Apply split panes with explicit split offsets, visible origin, and active pane.")
  record Split(
      int xSplitPosition,
      int ySplitPosition,
      int leftmostColumn,
      int topRow,
      ExcelPaneRegion activePane)
      implements PaneSpec {
    public Split {
      Objects.requireNonNull(activePane, "activePane must not be null");
      Validation.requireNonNegative(xSplitPosition, "xSplitPosition");
      Validation.requireNonNegative(ySplitPosition, "ySplitPosition");
      Validation.requireNonNegative(leftmostColumn, "leftmostColumn");
      Validation.requireNonNegative(topRow, "topRow");
      Validation.requireNonZeroPair(
          xSplitPosition, ySplitPosition, "xSplitPosition", "ySplitPosition");
      if (xSplitPosition == 0 && leftmostColumn != 0) {
        throw new IllegalArgumentException(
            "leftmostColumn must be 0 when xSplitPosition is 0: " + leftmostColumn);
      }
      if (ySplitPosition == 0 && topRow != 0) {
        throw new IllegalArgumentException("topRow must be 0 when ySplitPosition is 0: " + topRow);
      }
    }
  }

  /** Shared validation helpers for pane input constructors. */
  final class Validation {
    private Validation() {}

    static void requireNonNegative(int value, String fieldName) {
      if (value < 0) {
        throw new IllegalArgumentException(fieldName + " must not be negative");
      }
    }

    static void requireNonZeroPair(
        int firstValue, int secondValue, String firstFieldName, String secondFieldName) {
      if (firstValue == 0 && secondValue == 0) {
        throw new IllegalArgumentException(
            firstFieldName + " and " + secondFieldName + " must not both be 0");
      }
    }
  }
}
