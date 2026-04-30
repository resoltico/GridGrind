package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Shared range-rewrite helpers for data-validation authoring, readback, and health checks. */
final class ExcelDataValidationRangeSupport {
  private ExcelDataValidationRangeSupport() {}

  static void normalizeOverlappingSqref(XSSFSheet sheet, List<ExcelRange> cutouts) {
    if (!sheet.getCTWorksheet().isSetDataValidations()) {
      return;
    }
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    for (int validationIndex = dataValidations.sizeOfDataValidationArray() - 1;
        validationIndex >= 0;
        validationIndex--) {
      CTDataValidation validation = dataValidations.getDataValidationArray(validationIndex);
      List<String> retainedRanges = retainedRanges(validation, cutouts);
      if (retainedRanges.isEmpty()) {
        dataValidations.removeDataValidation(validationIndex);
        continue;
      }
      validation.setSqref(retainedRanges);
    }
    syncValidationCount(sheet);
  }

  static List<String> retainedRanges(CTDataValidation validation, List<ExcelRange> cutouts) {
    List<ExcelRange> retained =
        ExcelSqrefSupport.normalizedSqref(validation.getSqref()).stream()
            .map(ExcelRange::parse)
            .toList();
    List<ExcelRange> next = new ArrayList<>();
    for (ExcelRange cutout : cutouts) {
      next.clear();
      for (ExcelRange existingRange : retained) {
        next.addAll(subtract(existingRange, cutout));
      }
      retained = List.copyOf(next);
    }
    return retained.stream().map(ExcelDataValidationRangeSupport::formatRange).toList();
  }

  static List<ExcelRange> subtract(ExcelRange source, ExcelRange cutout) {
    if (!intersects(source, cutout)) {
      return List.of(source);
    }
    int intersectionFirstRow = Math.max(source.firstRow(), cutout.firstRow());
    int intersectionLastRow = Math.min(source.lastRow(), cutout.lastRow());
    int intersectionFirstColumn = Math.max(source.firstColumn(), cutout.firstColumn());
    int intersectionLastColumn = Math.min(source.lastColumn(), cutout.lastColumn());

    List<ExcelRange> retained = new ArrayList<>(4);
    if (source.firstRow() < intersectionFirstRow) {
      retained.add(
          new ExcelRange(
              source.firstRow(),
              intersectionFirstRow - 1,
              source.firstColumn(),
              source.lastColumn()));
    }
    if (intersectionLastRow < source.lastRow()) {
      retained.add(
          new ExcelRange(
              intersectionLastRow + 1,
              source.lastRow(),
              source.firstColumn(),
              source.lastColumn()));
    }
    if (source.firstColumn() < intersectionFirstColumn) {
      retained.add(
          new ExcelRange(
              intersectionFirstRow,
              intersectionLastRow,
              source.firstColumn(),
              intersectionFirstColumn - 1));
    }
    if (intersectionLastColumn < source.lastColumn()) {
      retained.add(
          new ExcelRange(
              intersectionFirstRow,
              intersectionLastRow,
              intersectionLastColumn + 1,
              source.lastColumn()));
    }
    return List.copyOf(retained);
  }

  static boolean intersects(ExcelRange first, ExcelRange second) {
    return first.firstRow() <= second.lastRow()
        && first.lastRow() >= second.firstRow()
        && first.firstColumn() <= second.lastColumn()
        && first.lastColumn() >= second.firstColumn();
  }

  static String formatRange(ExcelRange range) {
    return new CellRangeAddress(
            range.firstRow(), range.lastRow(), range.firstColumn(), range.lastColumn())
        .formatAsString();
  }

  static void syncValidationCount(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    int count = dataValidations.sizeOfDataValidationArray();
    if (count == 0) {
      sheet.getCTWorksheet().unsetDataValidations();
      return;
    }
    dataValidations.setCount(count);
  }

  static boolean matchesSelection(List<String> ranges, ExcelRangeSelection selection) {
    return switch (selection) {
      case ExcelRangeSelection.All _ -> true;
      case ExcelRangeSelection.Selected selected ->
          ranges.stream()
              .map(ExcelRange::parse)
              .anyMatch(
                  existing ->
                      selected.ranges().stream()
                          .map(ExcelRange::parse)
                          .anyMatch(selectedRange -> intersects(existing, selectedRange)));
    };
  }
}
