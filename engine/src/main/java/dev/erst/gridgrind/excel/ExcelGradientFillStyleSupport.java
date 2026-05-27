package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Map;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientStop;

/** Applies gradient and pattern fill mutations while reusing workbook fill ids. */
final class ExcelGradientFillStyleSupport {
  private final XSSFWorkbook workbook;
  private final StylesTableFillRegistryAccess fillRegistryAccess;
  private final Map<String, Integer> gradientFillIds;

  ExcelGradientFillStyleSupport(
      XSSFWorkbook workbook,
      StylesTableFillRegistryAccess fillRegistryAccess,
      Map<String, Integer> gradientFillIds) {
    this.workbook = workbook;
    this.fillRegistryAccess = fillRegistryAccess;
    this.gradientFillIds = gradientFillIds;
  }

  void indexExistingGradientFills() {
    List<XSSFCellFill> fills = fillsList();
    for (int fillId = 0; fillId < fills.size(); fillId++) {
      XSSFCellFill fill = fills.get(fillId);
      if (fill.getCTFill().isSetGradientFill()) {
        gradientFillIds.putIfAbsent(fill.getCTFill().xmlText(), fillId);
      }
    }
  }

  void applyGradientFillPatch(XSSFCellStyle cellStyle, ExcelGradientFill gradientPatch) {
    CTFill gradientFill = CTFill.Factory.newInstance();
    CTGradientFill gradient = gradientFill.addNewGradientFill();
    switch (gradientPatch) {
      case ExcelGradientFill.Path path -> {
        gradient.setType(
            org.openxmlformats.schemas.spreadsheetml.x2006.main.STGradientType.Enum.forString(
                "path"));
        path.left().ifPresent(gradient::setLeft);
        path.right().ifPresent(gradient::setRight);
        path.top().ifPresent(gradient::setTop);
        path.bottom().ifPresent(gradient::setBottom);
      }
      case ExcelGradientFill.Linear linear -> linear.degree().ifPresent(gradient::setDegree);
    }
    for (ExcelGradientStop stop : gradientPatch.stops()) {
      CTGradientStop ctStop = gradient.addNewStop();
      ctStop.setPosition(stop.position());
      ctStop.addNewColor().set(ExcelColorSupport.toXssfColor(workbook, stop.color()).getCTColor());
    }
    int gradientFillId = gradientFillId(gradientFill);
    cellStyle.getCoreXf().setApplyFill(true);
    cellStyle.getCoreXf().setFillId(gradientFillId);
  }

  void clearFillColors(XSSFCellStyle cellStyle) {
    cellStyle.setFillForegroundColor((XSSFColor) null);
    cellStyle.setFillBackgroundColor((XSSFColor) null);
  }

  private int gradientFillId(CTFill gradientFill) {
    String key = gradientFill.xmlText();
    Integer existingId = gradientFillIds.get(key);
    if (existingId != null) {
      return existingId;
    }
    int fillId =
        appendFill(new XSSFCellFill(gradientFill, workbook.getStylesSource().getIndexedColors()));
    gradientFillIds.put(key, fillId);
    return fillId;
  }

  private int appendFill(XSSFCellFill fill) {
    return fillRegistryAccess.appendFill(workbook.getStylesSource(), fill);
  }

  private List<XSSFCellFill> fillsList() {
    return fillRegistryAccess.fills(workbook.getStylesSource());
  }
}
