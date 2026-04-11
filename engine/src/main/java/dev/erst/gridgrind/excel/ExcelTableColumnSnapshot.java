package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual metadata for one persisted table column. */
public record ExcelTableColumnSnapshot(
    long id,
    String name,
    String uniqueName,
    String totalsRowLabel,
    String totalsRowFunction,
    String calculatedColumnFormula) {
  public ExcelTableColumnSnapshot {
    if (id < 0L) {
      throw new IllegalArgumentException("id must not be negative");
    }
    Objects.requireNonNull(name, "name must not be null");
    uniqueName = uniqueName == null ? "" : uniqueName;
    totalsRowLabel = totalsRowLabel == null ? "" : totalsRowLabel;
    totalsRowFunction = totalsRowFunction == null ? "" : totalsRowFunction;
    calculatedColumnFormula = calculatedColumnFormula == null ? "" : calculatedColumnFormula;
  }
}
