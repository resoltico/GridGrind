package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Inferred schema facts for one rectangular sheet window. */
public record SheetSchemaReport(
    String sheetName,
    String topLeftAddress,
    int rowCount,
    int columnCount,
    int dataRowCount,
    List<SchemaColumnReport> columns) {
  public SheetSchemaReport {
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
    if (dataRowCount < 0) {
      throw new IllegalArgumentException("dataRowCount must not be negative");
    }
    columns = GridGrindResponseSupport.copyValues(columns, "columns");
  }
}
