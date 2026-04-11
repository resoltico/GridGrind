package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual print-layout snapshot combining core and advanced page-setup state. */
public record ExcelPrintLayoutSnapshot(ExcelPrintLayout layout, ExcelPrintSetupSnapshot setup) {
  public ExcelPrintLayoutSnapshot {
    Objects.requireNonNull(layout, "layout must not be null");
    Objects.requireNonNull(setup, "setup must not be null");
  }
}
