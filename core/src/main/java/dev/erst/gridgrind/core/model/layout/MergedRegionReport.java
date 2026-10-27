package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import java.util.Objects;

/** One merged region captured from a sheet. */
public record MergedRegionReport(RangeLabel range) {
  public MergedRegionReport {
    Objects.requireNonNull(range, "range must not be null");
  }
}
