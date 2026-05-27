package dev.erst.gridgrind.excel;

import java.util.Optional;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetUpPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation;

/** Normalizes optional OOXML print-layout nodes after supported state changes. */
final class ExcelPrintLayoutNormalizationSupport {
  private ExcelPrintLayoutNormalizationSupport() {}

  static void normalizePrintNodes(XSSFSheet sheet) {
    normalizeHeaderFooterNode(sheet);
    normalizePageSetupNode(sheet);
    normalizePageSetupProperties(sheet);
    normalizePageMarginsNode(sheet);
  }

  static void normalizeHeaderFooterNode(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetHeaderFooter()) {
      return;
    }
    if (!ExcelPrintLayoutReadSupport.headerFooterText(sheet.getHeader()).isBlank()) {
      return;
    }
    if (!ExcelPrintLayoutReadSupport.headerFooterText(sheet.getFooter()).isBlank()) {
      return;
    }
    sheet.getCTWorksheet().unsetHeaderFooter();
  }

  static void normalizePageSetupNode(XSSFSheet sheet) {
    Optional<CTPageSetup> pageSetup = ExcelPrintLayoutReadSupport.pageSetup(sheet);
    if (pageSetup.isEmpty()) {
      return;
    }
    if (shouldUnsetPageSetupOrientation(sheet, pageSetup.orElseThrow())) {
      pageSetup.orElseThrow().unsetOrientation();
    }
    if (isEmptyPageSetup(pageSetup.orElseThrow())) {
      sheet.getCTWorksheet().unsetPageSetup();
    }
  }

  static boolean shouldUnsetPageSetupOrientation(XSSFSheet sheet, CTPageSetup pageSetup) {
    return pageSetup.isSetOrientation()
        && pageSetup.getOrientation() == STOrientation.PORTRAIT
        && !pageSetup.isSetFitToWidth()
        && !pageSetup.isSetFitToHeight();
  }

  static boolean isEmptyPageSetup(CTPageSetup pageSetup) {
    ExcelPrintSetup defaults = ExcelPrintSetup.defaults();
    return !pageSetup.isSetOrientation()
        && !pageSetup.isSetFitToWidth()
        && !pageSetup.isSetFitToHeight()
        && !pageSetup.isSetUsePrinterDefaults()
        && (!pageSetup.isSetPaperSize()
            || Math.toIntExact(pageSetup.getPaperSize()) == defaults.paperSize())
        && (!pageSetup.isSetDraft() || pageSetup.getDraft() == defaults.draft())
        && (!pageSetup.isSetBlackAndWhite()
            || pageSetup.getBlackAndWhite() == defaults.blackAndWhite())
        && (!pageSetup.isSetCopies() || Math.toIntExact(pageSetup.getCopies()) == defaults.copies())
        && (!pageSetup.isSetUseFirstPageNumber()
            || pageSetup.getUseFirstPageNumber() == defaults.useFirstPageNumber())
        && (!pageSetup.isSetFirstPageNumber()
            || Math.toIntExact(pageSetup.getFirstPageNumber()) == defaults.firstPageNumber());
  }

  static void normalizePageSetupProperties(XSSFSheet sheet) {
    Optional<CTSheetPr> sheetPr = ExcelPrintLayoutReadSupport.sheetPr(sheet);
    if (sheetPr.isEmpty()) {
      return;
    }
    Optional<CTPageSetUpPr> pageSetUpPr = ExcelPrintLayoutReadSupport.pageSetUpPr(sheet);
    if (pageSetUpPr.isPresent()
        && pageSetUpPr.orElseThrow().isSetFitToPage()
        && !pageSetUpPr.orElseThrow().getFitToPage()) {
      pageSetUpPr.orElseThrow().unsetFitToPage();
    }
    if (pageSetUpPr.isPresent() && !pageSetUpPr.orElseThrow().isSetFitToPage()) {
      sheetPr.orElseThrow().unsetPageSetUpPr();
    }
    if (!sheetPr.orElseThrow().isSetPageSetUpPr()) {
      sheet.getCTWorksheet().unsetSheetPr();
    }
  }

  static void normalizePageMarginsNode(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetPageMargins()) {
      return;
    }
    ExcelPrintMargins defaults = ExcelPrintSetup.defaults().margins();
    if (sheet.getMargin(PageMargin.LEFT) == defaults.left()
        && sheet.getMargin(PageMargin.RIGHT) == defaults.right()
        && sheet.getMargin(PageMargin.TOP) == defaults.top()
        && sheet.getMargin(PageMargin.BOTTOM) == defaults.bottom()
        && sheet.getMargin(PageMargin.HEADER) == defaults.header()
        && sheet.getMargin(PageMargin.FOOTER) == defaults.footer()) {
      sheet.getCTWorksheet().unsetPageMargins();
    }
  }
}
