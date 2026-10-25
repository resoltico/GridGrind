package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.List;
import java.util.Objects;

/** Inferred schema facts for one rectangular sheet window. */
public record SheetSchemaReport(
    SheetLabel sheetName,
    CellAddress topLeftAddress,
    int rowCount,
    int columnCount,
    int dataRowCount,
    List<SchemaColumnReport> columns) {
  public SheetSchemaReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be greater than 0");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be greater than 0");
    }
    if (dataRowCount < 0) {
      throw new IllegalArgumentException("dataRowCount must not be negative");
    }
    columns = WorkbookResultSupport.copyValues(columns, "columns");
  }
}
