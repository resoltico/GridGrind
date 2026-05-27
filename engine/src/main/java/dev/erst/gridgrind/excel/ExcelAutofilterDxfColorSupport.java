package dev.erst.gridgrind.excel;

import java.util.Optional;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;

/** Reads SpreadsheetML differential-style colors referenced by autofilter metadata. */
final class ExcelAutofilterDxfColorSupport {
  private ExcelAutofilterDxfColorSupport() {}

  static Optional<ExcelColorSnapshot> dxfColor(
      XSSFWorkbook workbook, long dxfId, boolean cellColor) {
    CTDxf dxf = dxfAt(workbook.getStylesSource(), dxfId).orElse(null);
    if (dxf == null) {
      return Optional.empty();
    }
    return preferredDxfColor(workbook, dxf, cellColor).or(() -> fallbackDxfColor(workbook, dxf));
  }

  private static Optional<ExcelColorSnapshot> preferredDxfColor(
      XSSFWorkbook workbook, CTDxf dxf, boolean cellColor) {
    return cellColor ? fillColor(workbook, dxf) : fontColor(workbook, dxf);
  }

  private static Optional<ExcelColorSnapshot> fallbackDxfColor(XSSFWorkbook workbook, CTDxf dxf) {
    return fillColor(workbook, dxf).or(() -> fontColor(workbook, dxf));
  }

  private static Optional<ExcelColorSnapshot> fillColor(XSSFWorkbook workbook, CTDxf dxf) {
    if (!dxf.isSetFill()
        || !dxf.getFill().isSetPatternFill()
        || !dxf.getFill().getPatternFill().isSetFgColor()) {
      return Optional.empty();
    }
    return ExcelColorSnapshotSupport.snapshot(
        workbook, dxf.getFill().getPatternFill().getFgColor());
  }

  private static Optional<ExcelColorSnapshot> fontColor(XSSFWorkbook workbook, CTDxf dxf) {
    if (!dxf.isSetFont() || dxf.getFont().sizeOfColorArray() == 0) {
      return Optional.empty();
    }
    return ExcelColorSnapshotSupport.snapshot(workbook, dxf.getFont().getColorArray(0));
  }

  static Optional<CTDxf> dxfAt(StylesTable stylesTable, long dxfId) {
    if (dxfId < 0L || dxfId >= stylesTable._getDXfsSize()) {
      return Optional.empty();
    }
    return Optional.of(stylesTable.getDxfAt(Math.toIntExact(dxfId)));
  }
}
