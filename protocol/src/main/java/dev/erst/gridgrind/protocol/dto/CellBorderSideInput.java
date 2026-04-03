package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Protocol-facing definition for one border side within a style patch. */
public record CellBorderSideInput(BorderStyle style) {
  public CellBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
  }
}
