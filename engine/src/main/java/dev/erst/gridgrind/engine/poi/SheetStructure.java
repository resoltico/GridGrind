package dev.erst.gridgrind.engine.poi;

import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;

/** Shared sheet-structure facts and range helpers used across workbook-core surfaces. */
@SuppressWarnings(
    "PMD.CommentRequired") // Shared structural operations have domain-specific, direct signatures.
public final class SheetStructure {
  private SheetStructure() {}

  public static void requireNoMergedRegionOverlap(Sheet sheet, CellBounds excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      CellRangeAddress existing = sheet.getMergedRegion(regionIndex);
      if (intersects(existing, excelRange)) {
        throw new IllegalArgumentException(
            "Merged range overlaps existing merged region: " + existing.formatAsString());
      }
    }
  }

  public static int findMergedRegionIndex(Sheet sheet, CellBounds excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      if (matches(sheet.getMergedRegion(regionIndex), excelRange)) {
        return regionIndex;
      }
    }
    return -1;
  }

  public static CellRangeAddress toCellRangeAddress(CellBounds excelRange) {
    return new CellRangeAddress(
        excelRange.firstRow(),
        excelRange.lastRow(),
        excelRange.firstColumn(),
        excelRange.lastColumn());
  }

  public static Optional<CellBounds> parseOptionalRange(String range) {
    try {
      return Optional.of(CellBounds.parse(range));
    } catch (IllegalArgumentException exception) {
      return Optional.empty();
    }
  }

  /** Formats one parsed workbook range back to canonical A1-style text. */
  public static String formatRange(CellBounds range) {
    return toCellRangeAddress(range).formatAsString();
  }

  public static boolean matches(CellRangeAddress rangeAddress, CellBounds excelRange) {
    return rangeAddress.getFirstRow() == excelRange.firstRow()
        && rangeAddress.getLastRow() == excelRange.lastRow()
        && rangeAddress.getFirstColumn() == excelRange.firstColumn()
        && rangeAddress.getLastColumn() == excelRange.lastColumn();
  }

  public static boolean intersects(CellRangeAddress rangeAddress, CellBounds excelRange) {
    return rangeAddress.getFirstRow() <= excelRange.lastRow()
        && rangeAddress.getLastRow() >= excelRange.firstRow()
        && rangeAddress.getFirstColumn() <= excelRange.lastColumn()
        && rangeAddress.getLastColumn() >= excelRange.firstColumn();
  }

  public static boolean intersects(CellBounds first, CellBounds second) {
    return first.firstRow() <= second.lastRow()
        && first.lastRow() >= second.firstRow()
        && first.firstColumn() <= second.lastColumn()
        && first.lastColumn() >= second.firstColumn();
  }

  public static boolean hasHeaderValue(Cell cell) {
    return !headerText(cell).isBlank();
  }

  public static boolean headerRowMissing(XSSFSheet sheet, CellBounds range) {
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

  public static int toColumnWidthUnits(double widthCharacters) {
    ExcelSheetLayoutLimits.requireColumnWidthCharacters(widthCharacters, "widthCharacters");
    int widthUnits = (int) Math.round(widthCharacters * 256.0d);
    return widthUnits;
  }

  public static float toRowHeightPoints(double heightPoints) {
    ExcelSheetLayoutLimits.requireRowHeightPoints(heightPoints, "heightPoints");
    return (float) heightPoints;
  }

  public static boolean shouldPreview(Cell cell) {
    return cell != null
        && shouldPreview(
            cell.getCellType(),
            cell.getCellStyle().getIndex(),
            cell.getHyperlink() != null,
            cell.getCellComment() != null);
  }

  public static boolean shouldPreview(
      CellType cellType, short styleIndex, boolean hasHyperlink, boolean hasComment) {
    return cellType != CellType.BLANK || styleIndex != 0 || hasHyperlink || hasComment;
  }

  /** Returns the normalized header texts contributed by one candidate header row range. */
  public static List<String> headerNames(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet, CellBounds range) {
    Row headerRow = sheet.getRow(range.firstRow());
    List<String> headers = new ArrayList<>(range.columnCount());
    for (int columnIndex = range.firstColumn(); columnIndex <= range.lastColumn(); columnIndex++) {
      headers.add(headerText(headerRow == null ? null : headerRow.getCell(columnIndex)));
    }
    return List.copyOf(headers);
  }

  /** Returns the normalized header text contributed by one candidate header cell. */
  public static String headerText(@Nullable Cell cell) {
    if (cell == null) {
      return "";
    }
    return switch (cell.getCellType()) {
      case BLANK -> "";
      case _NONE -> "";
      case STRING -> cell.getStringCellValue().trim();
      case FORMULA -> cell.getCellFormula().trim();
      default -> cell.toString().trim();
    };
  }
}
