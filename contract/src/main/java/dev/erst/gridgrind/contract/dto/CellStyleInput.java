package dev.erst.gridgrind.contract.dto;

/** Protocol-facing style patch used for range and cell presentation changes. */
public record CellStyleInput(
    String numberFormat,
    CellAlignmentInput alignment,
    CellFontInput font,
    CellFillInput fill,
    CellBorderInput border,
    CellProtectionInput protection) {
  public CellStyleInput {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    if (numberFormat == null
        && alignment == null
        && font == null
        && fill == null
        && border == null
        && protection == null) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }
}
