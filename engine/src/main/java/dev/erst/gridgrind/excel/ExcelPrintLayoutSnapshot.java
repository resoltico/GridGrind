package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import java.util.Objects;

/** Immutable factual print-layout snapshot combining core and advanced page-setup state. */
public record ExcelPrintLayoutSnapshot(ResolvedPrintLayout layout, PrintSetupReport setup) {
  public ExcelPrintLayoutSnapshot {
    Objects.requireNonNull(layout, "layout must not be null");
    Objects.requireNonNull(setup, "setup must not be null");
  }
}
