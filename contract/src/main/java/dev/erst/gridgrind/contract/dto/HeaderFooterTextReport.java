package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Plain left, center, and right header or footer text segments. */
public record HeaderFooterTextReport(String left, String center, String right) {
  public HeaderFooterTextReport {
    Objects.requireNonNull(left, "left must not be null");
    Objects.requireNonNull(center, "center must not be null");
    Objects.requireNonNull(right, "right must not be null");
  }
}
