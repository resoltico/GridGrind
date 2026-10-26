package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import java.util.List;
import java.util.Objects;

/** Cell read, write, preview, and array-formula operations for one sheet. */
public final class ExcelSheetCells {
  private final ExcelSheet sheet;
  private final ExcelSheetCellMutationSupport mutationSupport;
  private final ExcelSheetCellReadSupport readSupport;

  ExcelSheetCells(
      ExcelSheet sheet,
      ExcelSheetCellMutationSupport mutationSupport,
      ExcelSheetCellReadSupport readSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.mutationSupport =
        Objects.requireNonNull(mutationSupport, "mutationSupport must not be null");
    this.readSupport = Objects.requireNonNull(readSupport, "readSupport must not be null");
  }

  /** Writes a typed value to an A1-style address. */
  public ExcelSheetCells setCell(String address, ResolvedCellWrite value) {
    mutationSupport.setCell(address, value);
    return this;
  }

  /** Writes a rectangular matrix of values to an A1-style range such as {@code A1:C3}. */
  public ExcelSheetCells setRange(String range, List<List<ResolvedCellWrite>> rows) {
    mutationSupport.setRange(range, rows);
    return this;
  }

  /** Clears both contents and formatting in a rectangular A1-style range. */
  public ExcelSheetCells clearRange(String range) {
    mutationSupport.clearRange(range);
    return this;
  }

  /** Creates or replaces one dedicated array-formula group over a rectangular range. */
  public ExcelSheetCells setArrayFormula(String range, ResolvedArrayFormula formula) {
    mutationSupport.setArrayFormula(range, formula);
    return this;
  }

  /** Removes the array-formula group containing one addressed cell. */
  public ExcelSheetCells clearArrayFormula(String address) {
    mutationSupport.clearArrayFormula(address);
    return this;
  }

  /** Applies a style patch to every cell in a rectangular A1-style range. */
  public ExcelSheetCells applyStyle(String range, CellStylePatch style) {
    mutationSupport.applyStyle(range, style);
    return this;
  }

  /** Appends a new row using the next available row index. */
  public ExcelSheetCells appendRow(ResolvedCellWrite... values) {
    mutationSupport.appendRow(values);
    return this;
  }

  /** Reads a string cell by A1-style address. */
  public String text(String address) {
    return readSupport.text(address);
  }

  /** Reads a numeric cell by A1-style address, evaluating formulas when needed. */
  public double number(String address) {
    return readSupport.number(address);
  }

  /** Reads a boolean cell by A1-style address, evaluating formulas when needed. */
  public boolean bool(String address) {
    return readSupport.bool(address);
  }

  /** Reads the raw formula expression stored in a formula cell. */
  public String formula(String address) {
    return readSupport.formula(address);
  }

  /** Captures a formatted, typed snapshot of a single cell, returning blank for unwritten cells. */
  public CellReport snapshotCell(String address) {
    return readSupport.snapshotCell(address);
  }

  /** Captures one cell snapshot while evaluating formulas only when the projection needs it. */
  public CellReport snapshotCell(String address, CellReadProjection projection) {
    return readSupport.snapshotCell(address, projection);
  }

  /** Captures exact snapshots for the provided ordered A1 addresses. */
  public List<CellReport> snapshotCells(List<String> addresses) {
    return readSupport.snapshotCells(addresses);
  }

  /**
   * Captures ordered cell snapshots while evaluating formulas only when the projection needs it.
   */
  public List<CellReport> snapshotCells(List<String> addresses, CellReadProjection projection) {
    return readSupport.snapshotCells(addresses, projection);
  }

  /** Returns a compact preview of the top-left portion of the sheet. */
  public List<ExcelPreviewRow> preview(int maxRows, int maxColumns) {
    return readSupport.preview(maxRows, maxColumns, sheet.rows().lastIndex());
  }

  /** Returns an exact rectangular window of cell snapshots anchored at one top-left address. */
  public WorkbookSheetResult.Window window(String topLeftAddress, int rowCount, int columnCount) {
    return readSupport.window(topLeftAddress, rowCount, columnCount);
  }

  /**
   * Captures one rectangular window while evaluating formulas only when the projection needs it.
   */
  public WorkbookSheetResult.Window window(
      String topLeftAddress, int rowCount, int columnCount, CellReadProjection projection) {
    return readSupport.window(topLeftAddress, rowCount, columnCount, projection);
  }

  /** Returns factual array-formula groups on this sheet. */
  public List<ArrayFormulaReport> arrayFormulas() {
    return readSupport.arrayFormulas();
  }
}
