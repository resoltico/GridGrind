package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Rectangular window of cells anchored at one top-left address. */
public record WindowReport(
    String sheetName,
    String topLeftAddress,
    int rowCount,
    int columnCount,
    List<WindowRowReport> rows) {
  public WindowReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    if (topLeftAddress.isBlank()) {
      throw new IllegalArgumentException("topLeftAddress must not be blank");
    }
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be greater than 0");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be greater than 0");
    }
    rows = GridGrindResponseSupport.copyValues(rows, "rows");
  }
}
