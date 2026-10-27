package dev.erst.gridgrind.engine.layout;

import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelIndexDisplay;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import java.util.Objects;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Bound column editing and sizing operations. */
public final class ColumnAdapter {
  private final Sheet sheet;
  private final DataFormatter dataFormatter;

  /** Binds column editing and deterministic sizing to a sheet. */
  public ColumnAdapter(Sheet sheet, DataFormatter dataFormatter) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.dataFormatter = Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");
  }

  /** Sets the width of the inclusive column band in character units. */
  public ColumnAdapter setWidth(int firstColumnIndex, int lastColumnIndex, double widthCharacters) {
    requireNonNegative(firstColumnIndex, "firstColumnIndex");
    requireNonNegative(lastColumnIndex, "lastColumnIndex");
    requireOrderedSpan(firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");
    int widthUnits = SheetStructure.toColumnWidthUnits(widthCharacters);
    for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
      sheet.setColumnWidth(columnIndex, widthUnits);
    }
    return this;
  }

  /** Inserts blank columns with structural guards and adjacent visual formatting. */
  public ColumnAdapter insert(int columnIndex, int columnCount) {
    StructuralEdits.insertColumns(xssfSheet(), columnIndex, columnCount);
    return this;
  }

  /** Deletes the selected column band with structural guards. */
  public ColumnAdapter delete(ExcelColumnSpan columns) {
    Objects.requireNonNull(columns, "columns must not be null");
    StructuralEdits.deleteColumns(xssfSheet(), columns);
    return this;
  }

  /** Moves the selected column band by the signed delta. */
  public ColumnAdapter shift(ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(columns, "columns must not be null");
    StructuralEdits.shiftColumns(xssfSheet(), columns, delta);
    return this;
  }

  /** Sets the hidden state of the selected column band. */
  public ColumnAdapter setVisibility(ExcelColumnSpan columns, boolean hidden) {
    Objects.requireNonNull(columns, "columns must not be null");
    StructuralEdits.setColumnVisibility(xssfSheet(), columns, hidden);
    return this;
  }

  /** Groups the selected columns and applies the requested collapsed state. */
  public ColumnAdapter group(ExcelColumnSpan columns, boolean collapsed) {
    Objects.requireNonNull(columns, "columns must not be null");
    StructuralEdits.groupColumns(xssfSheet(), columns, collapsed);
    return this;
  }

  /** Removes outline grouping from the selected columns. */
  public ColumnAdapter ungroup(ExcelColumnSpan columns) {
    Objects.requireNonNull(columns, "columns must not be null");
    StructuralEdits.ungroupColumns(xssfSheet(), columns);
    return this;
  }

  /** Returns the last populated column index, or minus one for an empty sheet. */
  public int lastIndex() {
    return StructuralEdits.lastColumnIndex(xssfSheet());
  }

  /** Sizes populated columns deterministically from displayed content. */
  public ColumnAdapter autoSize() {
    DeterministicColumnSizer.autoSize(sheet, dataFormatter);
    return this;
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
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
