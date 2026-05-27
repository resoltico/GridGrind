package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Shared sheet-structure facts and range helpers used across workbook-core surfaces. */
public final class ExcelSheetStructureSupport {
  private ExcelSheetStructureSupport() {}

  static void requireNoMergedRegionOverlap(Sheet sheet, ExcelRange excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      CellRangeAddress existing = sheet.getMergedRegion(regionIndex);
      if (intersects(existing, excelRange)) {
        throw new IllegalArgumentException(
            "Merged range overlaps existing merged region: " + existing.formatAsString());
      }
    }
  }

  static int findMergedRegionIndex(Sheet sheet, ExcelRange excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      if (matches(sheet.getMergedRegion(regionIndex), excelRange)) {
        return regionIndex;
      }
    }
    return -1;
  }

  static CellRangeAddress toCellRangeAddress(ExcelRange excelRange) {
    return new CellRangeAddress(
        excelRange.firstRow(),
        excelRange.lastRow(),
        excelRange.firstColumn(),
        excelRange.lastColumn());
  }

  static Optional<ExcelRange> parseOptionalRange(String range) {
    try {
      return Optional.of(ExcelRange.parse(range));
    } catch (IllegalArgumentException exception) {
      return Optional.empty();
    }
  }

  /** Formats one parsed workbook range back to canonical A1-style text. */
  public static String formatRange(ExcelRange range) {
    return toCellRangeAddress(range).formatAsString();
  }

  static boolean matches(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return rangeAddress.getFirstRow() == excelRange.firstRow()
        && rangeAddress.getLastRow() == excelRange.lastRow()
        && rangeAddress.getFirstColumn() == excelRange.firstColumn()
        && rangeAddress.getLastColumn() == excelRange.lastColumn();
  }

  static boolean intersects(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return rangeAddress.getFirstRow() <= excelRange.lastRow()
        && rangeAddress.getLastRow() >= excelRange.firstRow()
        && rangeAddress.getFirstColumn() <= excelRange.lastColumn()
        && rangeAddress.getLastColumn() >= excelRange.firstColumn();
  }

  static boolean intersects(ExcelRange first, ExcelRange second) {
    return first.firstRow() <= second.lastRow()
        && first.lastRow() >= second.firstRow()
        && first.firstColumn() <= second.lastColumn()
        && first.lastColumn() >= second.firstColumn();
  }

  static boolean hasHeaderValue(Cell cell) {
    return !ExcelTableStructureSupport.headerText(cell).isBlank();
  }

  static boolean headerRowMissing(XSSFSheet sheet, ExcelRange range) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Row headerRow = sheet.getRow(range.firstRow());
    if (headerRow == null) {
      return true;
    }
    for (int columnIndex = range.firstColumn(); columnIndex <= range.lastColumn(); columnIndex++) {
      if (hasHeaderValue(headerRow.getCell(columnIndex))) {
        return false;
      }
    }
    return true;
  }

  static int toColumnWidthUnits(double widthCharacters) {
    ExcelSheetLayoutLimits.requireColumnWidthCharacters(widthCharacters, "widthCharacters");
    int widthUnits = (int) Math.round(widthCharacters * 256.0d);
    return widthUnits;
  }

  static float toRowHeightPoints(double heightPoints) {
    ExcelSheetLayoutLimits.requireRowHeightPoints(heightPoints, "heightPoints");
    return (float) heightPoints;
  }

  static boolean shouldPreview(Cell cell) {
    return cell != null
        && shouldPreview(
            cell.getCellType(),
            cell.getCellStyle().getIndex(),
            cell.getHyperlink() != null,
            cell.getCellComment() != null);
  }

  static boolean shouldPreview(
      CellType cellType, short styleIndex, boolean hasHyperlink, boolean hasComment) {
    return cellType != CellType.BLANK || styleIndex != 0 || hasHyperlink || hasComment;
  }
}
