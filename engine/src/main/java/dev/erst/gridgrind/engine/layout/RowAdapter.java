package dev.erst.gridgrind.engine.layout;

import dev.erst.gridgrind.core.vocabulary.ExcelIndexDisplay;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Bound row editing and sizing operations. */
public final class RowAdapter {
  private final Sheet sheet;

  /** Binds row editing and height operations to a sheet. */
  public RowAdapter(Sheet sheet) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
  }

  /** Sets the height of the inclusive row band in point units. */
  public RowAdapter setHeight(int firstRowIndex, int lastRowIndex, double heightPoints) {
    requireNonNegative(firstRowIndex, "firstRowIndex");
    requireNonNegative(lastRowIndex, "lastRowIndex");
    requireOrderedSpan(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");
    float heightPointsValue = SheetStructure.toRowHeightPoints(heightPoints);
    for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
      getOrCreateRow(rowIndex).setHeightInPoints(heightPointsValue);
    }
    return this;
  }

  /** Inserts blank rows with structural guards and adjacent visual formatting. */
  public RowAdapter insertRows(int rowIndex, int rowCount) {
    StructuralEdits.insertRows(xssfSheet(), rowIndex, rowCount);
    return this;
  }

  /** Deletes the selected row band with structural guards. */
  public RowAdapter deleteRows(ExcelRowSpan rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    StructuralEdits.deleteRows(xssfSheet(), rows);
    return this;
  }

  /** Moves the selected row band by the signed delta. */
  public RowAdapter shiftRows(ExcelRowSpan rows, int delta) {
    Objects.requireNonNull(rows, "rows must not be null");
    StructuralEdits.shiftRows(xssfSheet(), rows, delta);
    return this;
  }

  /** Sets the hidden state of the selected row band. */
  public RowAdapter setVisibility(ExcelRowSpan rows, boolean hidden) {
    Objects.requireNonNull(rows, "rows must not be null");
    StructuralEdits.setRowVisibility(xssfSheet(), rows, hidden);
    return this;
  }

  /** Groups the selected rows and applies the requested collapsed state. */
  public RowAdapter group(ExcelRowSpan rows, boolean collapsed) {
    Objects.requireNonNull(rows, "rows must not be null");
    StructuralEdits.groupRows(xssfSheet(), rows, collapsed);
    return this;
  }

  /** Removes outline grouping from the selected rows. */
  public RowAdapter ungroup(ExcelRowSpan rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    StructuralEdits.ungroupRows(xssfSheet(), rows);
    return this;
  }

  /** Returns the number of physically stored rows. */
  public int physicalCount() {
    return sheet.getPhysicalNumberOfRows();
  }

  /** Returns the last native row index. */
  public int lastIndex() {
    return sheet.getLastRowNum();
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private Row getOrCreateRow(int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(ExcelIndexDisplay.mustNotBeNegative(fieldName, value));
    }
  }

  private static void requireOrderedSpan(
      int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
    if (lastValue < firstValue) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeLessThan(
              lastFieldName, lastValue, firstFieldName, firstValue));
    }
  }
}
