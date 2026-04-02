package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;

/** Applies and reads sheet-view state such as panes and zoom. */
final class ExcelSheetViewSupport {
  private ExcelSheetViewSupport() {}

  /** Applies one explicit pane state to the provided sheet. */
  static void setPane(XSSFSheet sheet, ExcelSheetPane pane) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(pane, "pane must not be null");
    switch (pane) {
      case ExcelSheetPane.None _ -> sheet.createFreezePane(0, 0, 0, 0);
      case ExcelSheetPane.Frozen frozen ->
          sheet.createFreezePane(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case ExcelSheetPane.Split split ->
          sheet.createSplitPane(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane().toPoi());
    }
  }

  /** Returns the current pane state captured on the provided sheet. */
  static ExcelSheetPane pane(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    PaneInformation paneInformation = sheet.getPaneInformation();
    if (paneInformation == null) {
      return new ExcelSheetPane.None();
    }
    if (paneInformation.isFreezePane()) {
      return new ExcelSheetPane.Frozen(
          paneInformation.getVerticalSplitPosition(),
          paneInformation.getHorizontalSplitPosition(),
          paneInformation.getVerticalSplitLeftColumn(),
          paneInformation.getHorizontalSplitTopRow());
    }
    return new ExcelSheetPane.Split(
        paneInformation.getVerticalSplitPosition(),
        paneInformation.getHorizontalSplitPosition(),
        paneInformation.getVerticalSplitLeftColumn(),
        paneInformation.getHorizontalSplitTopRow(),
        ExcelPaneRegion.fromPoi(paneInformation.getActivePaneType()));
  }

  /** Applies one zoom percentage to the provided sheet. */
  static void setZoomPercent(XSSFSheet sheet, int zoomPercent) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireZoomPercent(zoomPercent);
    sheet.setZoom(zoomPercent);
  }

  /** Returns the effective zoom percentage for the provided sheet. */
  static int zoomPercent(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    CTSheetView sheetView = sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);
    return sheetView.isSetZoomScale() ? Math.toIntExact(sheetView.getZoomScale()) : 100;
  }

  /** Validates one requested zoom percentage. */
  static void requireZoomPercent(int zoomPercent) {
    if (zoomPercent < 10 || zoomPercent > 400) {
      throw new IllegalArgumentException(
          "zoomPercent must be between 10 and 400 inclusive: " + zoomPercent);
    }
  }
}
