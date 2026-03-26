package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelBorder;

/** Protocol-facing border patch used by {@link CellStyleInput}. */
public record CellBorderInput(
    CellBorderSideInput all,
    CellBorderSideInput top,
    CellBorderSideInput right,
    CellBorderSideInput bottom,
    CellBorderSideInput left) {
  public CellBorderInput {
    if (all == null && top == null && right == null && bottom == null && left == null) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }

  /** Converts this transport model into the engine border model. */
  public ExcelBorder toExcelBorder() {
    return new ExcelBorder(
        all == null ? null : all.toExcelBorderSide(),
        top == null ? null : top.toExcelBorderSide(),
        right == null ? null : right.toExcelBorderSide(),
        bottom == null ? null : bottom.toExcelBorderSide(),
        left == null ? null : left.toExcelBorderSide());
  }
}
