package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Factual metadata for one persisted workbook table column. */
public record TableColumnReport(
    long id,
    String name,
    String uniqueName,
    String totalsRowLabel,
    String totalsRowFunction,
    String calculatedColumnFormula) {
  public TableColumnReport {
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
