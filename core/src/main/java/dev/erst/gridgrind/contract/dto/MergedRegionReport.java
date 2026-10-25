package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** One merged region captured from a sheet. */
public record MergedRegionReport(String range) {
  public MergedRegionReport {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
  }
}
