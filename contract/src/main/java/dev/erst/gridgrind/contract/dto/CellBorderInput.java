package dev.erst.gridgrind.contract.dto;

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
}
