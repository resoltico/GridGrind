package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.stream.IntStream;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HeaderFooter;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetUpPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetPr;

/** Applies and reads supported print-layout state for one sheet. */
final class ExcelPrintLayoutController {
  /** Applies the provided print layout as the authoritative supported print state. */
  void setPrintLayout(XSSFSheet sheet, ExcelPrintLayout layout) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(layout, "layout must not be null");

    applyPrintArea(sheet, layout.printArea());
    applyOrientation(sheet, layout.orientation());
    applyScaling(sheet, layout.scaling());
    applyRepeatingRows(sheet, layout.repeatingRows());
    applyRepeatingColumns(sheet, layout.repeatingColumns());
    applyHeader(sheet.getHeader(), layout.header());
    applyFooter(sheet.getFooter(), layout.footer());
    normalizePrintNodes(sheet);
  }

  /** Clears the supported print layout state from the provided sheet. */
  void clearPrintLayout(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    setPrintLayout(sheet, ExcelPrintLayout.defaults());
  }

  /** Returns the supported print layout state currently stored for the provided sheet. */
  ExcelPrintLayout printLayout(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return new ExcelPrintLayout(
        printArea(sheet),
        orientation(sheet),
        scaling(sheet),
        repeatingRows(sheet),
        repeatingColumns(sheet),
        headerFooterText(sheet.getHeader()),
        headerFooterText(sheet.getFooter()));
  }

  /** Returns the full factual print-layout snapshot currently stored for the provided sheet. */
  ExcelPrintLayoutSnapshot printLayoutSnapshot(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return new ExcelPrintLayoutSnapshot(
        printLayout(sheet),
        new ExcelPrintSetupSnapshot(
            new ExcelPrintMarginsSnapshot(
                sheet.getMargin(XSSFSheet.LeftMargin),
                sheet.getMargin(XSSFSheet.RightMargin),
                sheet.getMargin(XSSFSheet.TopMargin),
                sheet.getMargin(XSSFSheet.BottomMargin),
                sheet.getMargin(XSSFSheet.HeaderMargin),
                sheet.getMargin(XSSFSheet.FooterMargin)),
            sheet.getHorizontallyCenter(),
            sheet.getVerticallyCenter(),
            sheet.getPrintSetup().getPaperSize(),
            sheet.getPrintSetup().getDraft(),
            sheet.getPrintSetup().getNoColor(),
            sheet.getPrintSetup().getCopies(),
            sheet.getPrintSetup().getUsePage(),
            sheet.getPrintSetup().getPageStart(),
            IntStream.of(sheet.getRowBreaks()).boxed().toList(),
            IntStream.of(sheet.getColumnBreaks()).boxed().toList()));
  }

  private static void applyPrintArea(XSSFSheet sheet, ExcelPrintLayout.Area printArea) {
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    switch (printArea) {
      case ExcelPrintLayout.Area.None _ -> sheet.getWorkbook().removePrintArea(sheetIndex);
      case ExcelPrintLayout.Area.Range range -> {
        ExcelRange parsed = ExcelRange.parse(range.range());
        sheet
            .getWorkbook()
            .setPrintArea(
                sheetIndex,
                parsed.firstColumn(),
                parsed.lastColumn(),
                parsed.firstRow(),
                parsed.lastRow());
      }
    }
  }

  private static void applyOrientation(XSSFSheet sheet, ExcelPrintOrientation orientation) {
    sheet.getPrintSetup().setLandscape(orientation == ExcelPrintOrientation.LANDSCAPE);
  }

  static void applyScaling(XSSFSheet sheet, ExcelPrintLayout.Scaling scaling) {
    XSSFSheet sheetRef = Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFPrintSetup printSetup = sheetRef.getPrintSetup();
    switch (scaling) {
      case ExcelPrintLayout.Scaling.Automatic _ -> {
        sheetRef.setFitToPage(false);
        CTPageSetup pageSetup =
            Objects.requireNonNull(
                pageSetupOrNull(sheetRef), "page setup must exist after fit-to-page toggle");
        if (pageSetup.isSetFitToWidth()) {
          pageSetup.unsetFitToWidth();
        }
        if (pageSetup.isSetFitToHeight()) {
          pageSetup.unsetFitToHeight();
        }
      }
      case ExcelPrintLayout.Scaling.Fit fit -> {
        sheetRef.setFitToPage(true);
        printSetup.setFitWidth((short) fit.widthPages());
        printSetup.setFitHeight((short) fit.heightPages());
      }
    }
  }

  private static void applyRepeatingRows(XSSFSheet sheet, ExcelPrintLayout.TitleRows titleRows) {
    switch (titleRows) {
      case ExcelPrintLayout.TitleRows.None _ -> sheet.setRepeatingRows(null);
      case ExcelPrintLayout.TitleRows.Band band ->
          sheet.setRepeatingRows(
              new CellRangeAddress(band.firstRowIndex(), band.lastRowIndex(), -1, -1));
    }
  }

  private static void applyRepeatingColumns(
      XSSFSheet sheet, ExcelPrintLayout.TitleColumns titleColumns) {
    switch (titleColumns) {
      case ExcelPrintLayout.TitleColumns.None _ -> sheet.setRepeatingColumns(null);
      case ExcelPrintLayout.TitleColumns.Band band ->
          sheet.setRepeatingColumns(
              new CellRangeAddress(-1, -1, band.firstColumnIndex(), band.lastColumnIndex()));
    }
  }

  private static void applyHeader(Header header, ExcelHeaderFooterText text) {
    header.setLeft(text.left());
    header.setCenter(text.center());
    header.setRight(text.right());
  }

  private static void applyFooter(Footer footer, ExcelHeaderFooterText text) {
    footer.setLeft(text.left());
    footer.setCenter(text.center());
    footer.setRight(text.right());
  }

  private static ExcelPrintLayout.Area printArea(XSSFSheet sheet) {
    String printArea = storedPrintAreaFormulaOrNull(sheet);
    if (printArea == null || printArea.isBlank()) {
      return new ExcelPrintLayout.Area.None();
    }
    AreaReference areaReference = new AreaReference(printArea, SpreadsheetVersion.EXCEL2007);
    return new ExcelPrintLayout.Area.Range(
        new CellRangeAddress(
                areaReference.getFirstCell().getRow(),
                areaReference.getLastCell().getRow(),
                areaReference.getFirstCell().getCol(),
                areaReference.getLastCell().getCol())
            .formatAsString());
  }

  static String storedPrintAreaFormulaOrNull(XSSFSheet sheet) {
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    String rawPrintArea = rawPrintAreaFormulaOrNull(sheet.getWorkbook(), sheetIndex);
    if (rawPrintArea == null || rawPrintArea.isBlank()) {
      return rawPrintArea;
    }
    return sheet.getWorkbook().getPrintArea(sheetIndex);
  }

  static String rawPrintAreaFormulaOrNull(XSSFWorkbook workbook, int sheetIndex) {
    if (!workbook.getCTWorkbook().isSetDefinedNames()) {
      return null;
    }
    for (CTDefinedName definedName :
        workbook.getCTWorkbook().getDefinedNames().getDefinedNameList()) {
      if (XSSFName.BUILTIN_PRINT_AREA.equals(definedName.getName())
          && definedName.isSetLocalSheetId()
          && definedName.getLocalSheetId() == sheetIndex) {
        return definedName.getStringValue();
      }
    }
    return null;
  }

  private static ExcelPrintOrientation orientation(XSSFSheet sheet) {
    return sheet.getPrintSetup().getLandscape()
        ? ExcelPrintOrientation.LANDSCAPE
        : ExcelPrintOrientation.PORTRAIT;
  }

  static ExcelPrintLayout.Scaling scaling(XSSFSheet sheet) {
    CTPageSetUpPr pageSetUpPr = pageSetUpPrOrNull(sheet);
    if (pageSetUpPr == null || !pageSetUpPr.isSetFitToPage() || !pageSetUpPr.getFitToPage()) {
      return new ExcelPrintLayout.Scaling.Automatic();
    }
    XSSFPrintSetup printSetup = sheet.getPrintSetup();
    return new ExcelPrintLayout.Scaling.Fit(printSetup.getFitWidth(), printSetup.getFitHeight());
  }

  static ExcelPrintLayout.TitleRows repeatingRows(XSSFSheet sheet) {
    CellRangeAddress repeatingRows = sheet.getRepeatingRows();
    if (repeatingRows == null) {
      return new ExcelPrintLayout.TitleRows.None();
    }
    return new ExcelPrintLayout.TitleRows.Band(
        repeatingRows.getFirstRow(), repeatingRows.getLastRow());
  }

  static ExcelPrintLayout.TitleColumns repeatingColumns(XSSFSheet sheet) {
    CellRangeAddress repeatingColumns = sheet.getRepeatingColumns();
    if (repeatingColumns == null) {
      return new ExcelPrintLayout.TitleColumns.None();
    }
    return new ExcelPrintLayout.TitleColumns.Band(
        repeatingColumns.getFirstColumn(), repeatingColumns.getLastColumn());
  }

  private static ExcelHeaderFooterText headerFooterText(HeaderFooter headerFooter) {
    return new ExcelHeaderFooterText(
        Objects.toString(headerFooter.getLeft(), ""),
        Objects.toString(headerFooter.getCenter(), ""),
        Objects.toString(headerFooter.getRight(), ""));
  }

  private static void normalizePrintNodes(XSSFSheet sheet) {
    normalizeHeaderFooterNode(sheet);
    normalizePageSetupNode(sheet);
    normalizePageSetupProperties(sheet);
  }

  static void normalizeHeaderFooterNode(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetHeaderFooter()) {
      return;
    }
    if (!headerFooterText(sheet.getHeader()).isBlank()) {
      return;
    }
    if (!headerFooterText(sheet.getFooter()).isBlank()) {
      return;
    }
    sheet.getCTWorksheet().unsetHeaderFooter();
  }

  static void normalizePageSetupNode(XSSFSheet sheet) {
    CTPageSetup pageSetup = pageSetupOrNull(sheet);
    if (pageSetup == null) {
      return;
    }
    if (shouldUnsetPageSetupOrientation(sheet, pageSetup)) {
      pageSetup.unsetOrientation();
    }
    if (isEmptyPageSetup(pageSetup)) {
      sheet.getCTWorksheet().unsetPageSetup();
    }
  }

  static boolean shouldUnsetPageSetupOrientation(XSSFSheet sheet, CTPageSetup pageSetup) {
    return pageSetup.isSetOrientation()
        && !sheet.getPrintSetup().getLandscape()
        && !pageSetup.isSetFitToWidth()
        && !pageSetup.isSetFitToHeight();
  }

  static boolean isEmptyPageSetup(CTPageSetup pageSetup) {
    return !pageSetup.isSetOrientation()
        && !pageSetup.isSetFitToWidth()
        && !pageSetup.isSetFitToHeight()
        && !pageSetup.isSetUsePrinterDefaults();
  }

  static void normalizePageSetupProperties(XSSFSheet sheet) {
    CTSheetPr sheetPr = sheetPrOrNull(sheet);
    if (sheetPr == null) {
      return;
    }
    CTPageSetUpPr pageSetUpPr = pageSetUpPrOrNull(sheet);
    if (pageSetUpPr != null && pageSetUpPr.isSetFitToPage() && !pageSetUpPr.getFitToPage()) {
      pageSetUpPr.unsetFitToPage();
    }
    if (pageSetUpPr != null && !pageSetUpPr.isSetFitToPage()) {
      sheetPr.unsetPageSetUpPr();
    }
    if (!sheetPr.isSetPageSetUpPr()) {
      sheet.getCTWorksheet().unsetSheetPr();
    }
  }

  static CTPageSetup pageSetupOrNull(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetPageSetup() ? sheet.getCTWorksheet().getPageSetup() : null;
  }

  static CTSheetPr sheetPrOrNull(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetSheetPr() ? sheet.getCTWorksheet().getSheetPr() : null;
  }

  static CTPageSetUpPr pageSetUpPrOrNull(XSSFSheet sheet) {
    CTSheetPr sheetPr = sheetPrOrNull(sheet);
    if (sheetPr == null || !sheetPr.isSetPageSetUpPr()) {
      return null;
    }
    return sheetPr.getPageSetUpPr();
  }
}
