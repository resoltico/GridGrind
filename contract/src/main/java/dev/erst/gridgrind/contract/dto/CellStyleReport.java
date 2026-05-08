package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Effective style facts returned with every analyzed cell. */
public record CellStyleReport(
    String numberFormat,
    CellAlignmentReport alignment,
    CellFontReport font,
    CellFillReport fill,
    CellBorderReport border,
    CellProtectionReport protection) {
  public CellStyleReport {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    Objects.requireNonNull(alignment, "alignment must not be null");
    Objects.requireNonNull(font, "font must not be null");
    Objects.requireNonNull(fill, "fill must not be null");
    Objects.requireNonNull(border, "border must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
  }
}
