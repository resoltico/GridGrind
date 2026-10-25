package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Effective border facts reported with every analyzed cell. */
public record CellBorderReport(
    CellBorderSideReport top,
    CellBorderSideReport right,
    CellBorderSideReport bottom,
    CellBorderSideReport left) {
  public CellBorderReport {
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
  }
}
