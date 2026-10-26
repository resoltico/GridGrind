package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import java.util.Objects;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;

/** Applies and reads sheet-view state such as panes and zoom. */
final class ExcelSheetViewSupport {
  private ExcelSheetViewSupport() {}

  /** Applies one explicit pane state to the provided sheet. */
  static void setPane(XSSFSheet sheet, PaneSpec pane) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(pane, "pane must not be null");
    switch (pane) {
      case PaneSpec.None _ -> sheet.createFreezePane(0, 0, 0, 0);
      case PaneSpec.Frozen frozen ->
          sheet.createFreezePane(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case PaneSpec.Split split ->
          sheet.createSplitPane(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              ExcelPanePoiBridge.toPoi(split.activePane()));
    }
  }

  /** Returns the current pane state captured on the provided sheet. */
  static dev.erst.gridgrind.core.model.layout.PaneReport pane(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    PaneInformation paneInformation = sheet.getPaneInformation();
    if (paneInformation == null) {
      return new dev.erst.gridgrind.core.model.layout.PaneReport.None();
    }
    if (paneInformation.isFreezePane()) {
      return new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(
          paneInformation.getVerticalSplitPosition(),
          paneInformation.getHorizontalSplitPosition(),
          paneInformation.getVerticalSplitLeftColumn(),
          paneInformation.getHorizontalSplitTopRow());
    }
    return new dev.erst.gridgrind.core.model.layout.PaneReport.Split(
        paneInformation.getVerticalSplitPosition(),
        paneInformation.getHorizontalSplitPosition(),
        paneInformation.getVerticalSplitLeftColumn(),
        paneInformation.getHorizontalSplitTopRow(),
        ExcelPanePoiBridge.fromPoi(paneInformation.getActivePaneType()));
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
  static void requireZoomPercent(int zoomPercent) { // LIM-022
    ExcelSheetLayoutLimits.requireZoomPercent(zoomPercent, "zoomPercent");
  }
}
