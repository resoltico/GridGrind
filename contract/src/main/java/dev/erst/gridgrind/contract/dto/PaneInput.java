package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import java.util.Objects;

/** One explicit sheet pane state requested through the wire protocol. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PaneInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = PaneInput.Frozen.class, name = "FROZEN"),
  @JsonSubTypes.Type(value = PaneInput.Split.class, name = "SPLIT")
})
public sealed interface PaneInput permits PaneInput.None, PaneInput.Frozen, PaneInput.Split {
  /** Sheet has no active pane split or freeze state. */
  record None() implements PaneInput {}

  /** Sheet is frozen at the provided split and visible-origin coordinates. */
  record Frozen(Integer splitColumn, Integer splitRow, Integer leftmostColumn, Integer topRow)
      implements PaneInput {
    public Frozen {
      Objects.requireNonNull(splitColumn, "splitColumn must not be null");
      Objects.requireNonNull(splitRow, "splitRow must not be null");
      Objects.requireNonNull(leftmostColumn, "leftmostColumn must not be null");
      Objects.requireNonNull(topRow, "topRow must not be null");
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
  record Split(
      Integer xSplitPosition,
      Integer ySplitPosition,
      Integer leftmostColumn,
      Integer topRow,
      ExcelPaneRegion activePane)
      implements PaneInput {
    public Split {
      Objects.requireNonNull(xSplitPosition, "xSplitPosition must not be null");
      Objects.requireNonNull(ySplitPosition, "ySplitPosition must not be null");
      Objects.requireNonNull(leftmostColumn, "leftmostColumn must not be null");
      Objects.requireNonNull(topRow, "topRow must not be null");
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
