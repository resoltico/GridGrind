package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;

/** Owns explicit column-definition normalization, shifting, and readback on XSSF sheets. */
final class ExcelColumnDefinitionSupport {
  private ExcelColumnDefinitionSupport() {}

  /** Returns the last column index implied by cells or explicit column metadata. */
  static int lastColumnIndex(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int lastColumnIndex = -1;
    for (Row row : sheet) {
      lastColumnIndex = Math.max(lastColumnIndex, row.getLastCellNum() - 1);
    }
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        lastColumnIndex = Math.max(lastColumnIndex, (int) col.getMax() - 1);
      }
    }
    return lastColumnIndex;
  }

  /** Returns the sheet column layouts including hidden, outline, and collapsed state. */
  static List<WorkbookReadResult.ColumnLayout> columnLayouts(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      return List.of();
    }
    Map<Integer, CTCol> effectiveColumns = effectiveColumnDefinitions(sheet);
    List<WorkbookReadResult.ColumnLayout> columns = new ArrayList<>(lastColumnIndex + 1);
    for (int columnIndex = 0; columnIndex <= lastColumnIndex; columnIndex++) {
      CTCol columnDefinition = effectiveColumns.get(columnIndex);
      columns.add(
          new WorkbookReadResult.ColumnLayout(
              columnIndex,
              sheet.getColumnWidth(columnIndex) / 256.0d,
              columnDefinition != null && columnDefinition.getHidden(),
              columnDefinition == null ? 0 : (int) columnDefinition.getOutlineLevel(),
              columnDefinition != null && columnDefinition.getCollapsed()));
    }
    return List.copyOf(columns);
  }

  /** Applies a collapsed marker to the explicit column definition that owns the target index. */
  static void setColumnCollapsed(XSSFSheet sheet, int columnIndex, boolean collapsed) {
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 >= col.getMin() && columnIndex + 1 <= col.getMax()) {
          col.setCollapsed(collapsed);
        }
      }
    }
  }

  static void normalizeColumnDefinitionContainer(XSSFSheet sheet) {
    if (!requiresColumnDefinitionCanonicalization(sheet)) {
      return;
    }
    canonicalizeColumnDefinitions(sheet);
  }

  static void canonicalizeColumnDefinitions(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    rebuildColumnDefinitions(sheet, effectiveColumnDefinitions(sheet));
  }

  static void rebuildColumnDefinitions(XSSFSheet sheet, Map<Integer, CTCol> explicitColumns) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(explicitColumns, "explicitColumns must not be null");
    rebuildColumnDefinitionsInternal(sheet, explicitColumns);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Map<Integer, CTCol> snapshotColumnDefinitions(XSSFSheet sheet) {
    Map<Integer, CTCol> explicitColumns = new LinkedHashMap<>();
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        for (int columnIndex = (int) col.getMin() - 1;
            columnIndex <= (int) col.getMax() - 1;
            columnIndex++) {
          CTCol definition = copyOf(col);
          definition.setMin(columnIndex + 1L);
          definition.setMax(columnIndex + 1L);
          explicitColumns.put(columnIndex, definition);
        }
      }
    }
    return Map.copyOf(explicitColumns);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Map<Integer, CTCol> shiftedForInsert(
      Map<Integer, CTCol> explicitColumns, int columnIndex, int columnCount) {
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) ->
            shifted.put(
                existingColumnIndex >= columnIndex
                    ? existingColumnIndex + columnCount
                    : existingColumnIndex,
                definition));
    return Map.copyOf(shifted);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Map<Integer, CTCol> shiftedForDelete(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns) {
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          if (existingColumnIndex < columns.firstColumnIndex()) {
            shifted.put(existingColumnIndex, definition);
            return;
          }
          if (existingColumnIndex > columns.lastColumnIndex()) {
            shifted.put(existingColumnIndex - columns.count(), definition);
          }
        });
    return Map.copyOf(shifted);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Map<Integer, CTCol> shiftedForShift(
      Map<Integer, CTCol> explicitColumns, ExcelColumnSpan columns, int delta) {
    int destinationFirstColumn = columns.firstColumnIndex() + delta;
    int destinationLastColumn = columns.lastColumnIndex() + delta;
    int overwrittenFirstColumn = Math.min(destinationFirstColumn, destinationLastColumn);
    int overwrittenLastColumn = Math.max(destinationFirstColumn, destinationLastColumn);
    Map<Integer, CTCol> shifted = new LinkedHashMap<>();
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          boolean inSource =
              existingColumnIndex >= columns.firstColumnIndex()
                  && existingColumnIndex <= columns.lastColumnIndex();
          boolean overwrittenDestination =
              existingColumnIndex >= overwrittenFirstColumn
                  && existingColumnIndex <= overwrittenLastColumn;
          if (!inSource && !overwrittenDestination) {
            shifted.put(existingColumnIndex, definition);
          }
        });
    explicitColumns.forEach(
        (existingColumnIndex, definition) -> {
          if (existingColumnIndex >= columns.firstColumnIndex()
              && existingColumnIndex <= columns.lastColumnIndex()) {
            shifted.put(existingColumnIndex + delta, definition);
          }
        });
    return Map.copyOf(shifted);
  }

  private static void rebuildColumnDefinitionsInternal(
      XSSFSheet sheet, Map<Integer, CTCol> explicitColumns) {
    CTCols cols = CTCols.Factory.newInstance();
    explicitColumns.entrySet().stream()
        .sorted(Map.Entry.comparingByKey())
        .forEach(
            entry -> {
              CTCol definition = copyOf(entry.getValue());
              definition.setMin(entry.getKey() + 1L);
              definition.setMax(entry.getKey() + 1L);
              cols.addNewCol().set(definition);
            });
    sheet.getCTWorksheet().setColsArray(new CTCols[] {cols});
  }

  private static boolean requiresColumnDefinitionCanonicalization(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().sizeOfColsArray() != 1) {
      return true;
    }
    boolean[] seenColumns = new boolean[ExcelColumnSpan.MAX_COLUMN_INDEX + 1];
    for (CTCol col : sheet.getCTWorksheet().getColsArray(0).getColList()) {
      if (col.getMin() != col.getMax()) {
        return true;
      }
      if (isSemanticallyEmptyColumnDefinition(col)) {
        return true;
      }
      for (int columnIndex = (int) col.getMin() - 1;
          columnIndex <= (int) col.getMax() - 1;
          columnIndex++) {
        if (seenColumns[columnIndex]) {
          return true;
        }
        seenColumns[columnIndex] = true;
      }
    }
    return false;
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<Integer, CTCol> effectiveColumnDefinitions(XSSFSheet sheet) {
    Map<Integer, CTCol> explicitColumns = snapshotColumnDefinitions(sheet);
    int lastColumnIndex = lastColumnIndex(sheet);
    if (lastColumnIndex < 0) {
      return Map.of();
    }
    Map<Integer, CTCol> effectiveColumns = new LinkedHashMap<>();
    for (int columnIndex = 0; columnIndex <= lastColumnIndex; columnIndex++) {
      CTCol effectiveDefinition =
          effectiveColumnDefinition(sheet, columnIndex, explicitColumns.get(columnIndex));
      if (effectiveDefinition != null) {
        effectiveColumns.put(columnIndex, effectiveDefinition);
      }
    }
    return Map.copyOf(effectiveColumns);
  }

  private static CTCol effectiveColumnDefinition(
      XSSFSheet sheet, int columnIndex, CTCol baseDefinition) {
    boolean hidden = false;
    long outlineLevel = 0L;
    boolean collapsed = false;
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 < col.getMin() || columnIndex + 1 > col.getMax()) {
          continue;
        }
        hidden |= col.getHidden();
        outlineLevel = Math.max(outlineLevel, col.getOutlineLevel());
        collapsed |= col.getCollapsed();
      }
    }
    if (!hasMeaningfulColumnDefinition(baseDefinition, hidden, outlineLevel, collapsed)) {
      return null;
    }
    CTCol effectiveDefinition = copyOf(baseDefinition);
    effectiveDefinition.setMin(columnIndex + 1L);
    effectiveDefinition.setMax(columnIndex + 1L);
    effectiveDefinition.setHidden(hidden);
    effectiveDefinition.setOutlineLevel((short) outlineLevel);
    effectiveDefinition.setCollapsed(collapsed);
    return effectiveDefinition;
  }

  private static boolean hasMeaningfulColumnDefinition(
      CTCol baseDefinition, boolean hidden, long outlineLevel, boolean collapsed) {
    return hidden
        || outlineLevel > 0L
        || collapsed
        || (baseDefinition != null && !isSemanticallyEmptyColumnDefinition(baseDefinition));
  }

  private static boolean isSemanticallyEmptyColumnDefinition(CTCol definition) {
    return !definition.getHidden()
        && definition.getOutlineLevel() == 0
        && !definition.getCollapsed()
        && !definition.getCustomWidth()
        && !definition.getBestFit()
        && !definition.getPhonetic()
        && (!definition.isSetStyle() || definition.getStyle() == 0L);
  }

  private static CTCol copyOf(CTCol original) {
    CTCol copy = CTCol.Factory.newInstance();
    copy.set(original);
    return copy;
  }
}
