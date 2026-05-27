package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
    applySetup(sheet, layout.setup());
    ExcelPrintLayoutNormalizationSupport.normalizePrintNodes(sheet);
  }

  /** Clears the supported print layout state from the provided sheet. */
  void clearPrintLayout(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    applyPrintArea(sheet, new ExcelPrintLayout.Area.None());
    applyRepeatingRows(sheet, new ExcelPrintLayout.TitleRows.None());
    applyRepeatingColumns(sheet, new ExcelPrintLayout.TitleColumns.None());
    applyHeader(sheet.getHeader(), ExcelHeaderFooterText.blank());
    applyFooter(sheet.getFooter(), ExcelHeaderFooterText.blank());
    replaceBreaks(
        sheet.getRowBreaks(), java.util.List.of(), sheet::removeRowBreak, sheet::setRowBreak);
    replaceBreaks(
        sheet.getColumnBreaks(),
        java.util.List.of(),
        sheet::removeColumnBreak,
        sheet::setColumnBreak);
    sheet.setPrintGridlines(false);
    sheet.setHorizontallyCenter(false);
    sheet.setVerticallyCenter(false);
    sheet.setFitToPage(false);
    if (sheet.getCTWorksheet().isSetPageSetup()) {
      sheet.getCTWorksheet().unsetPageSetup();
    }
    if (sheet.getCTWorksheet().isSetPageMargins()) {
      sheet.getCTWorksheet().unsetPageMargins();
    }
    ExcelPrintLayoutNormalizationSupport.normalizePrintNodes(sheet);
  }

  /** Returns the supported print layout state currently stored for the provided sheet. */
  ExcelPrintLayout printLayout(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return new ExcelPrintLayout(
        ExcelPrintLayoutReadSupport.printArea(sheet),
        ExcelPrintLayoutReadSupport.orientation(sheet),
        ExcelPrintLayoutReadSupport.scaling(sheet),
        ExcelPrintLayoutReadSupport.repeatingRows(sheet),
        ExcelPrintLayoutReadSupport.repeatingColumns(sheet),
        ExcelPrintLayoutReadSupport.headerFooterText(sheet.getHeader()),
        ExcelPrintLayoutReadSupport.headerFooterText(sheet.getFooter()),
        ExcelPrintLayoutReadSupport.setup(sheet));
  }

  /** Returns the full factual print-layout snapshot currently stored for the provided sheet. */
  ExcelPrintLayoutSnapshot printLayoutSnapshot(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelPrintSetup setup = ExcelPrintLayoutReadSupport.setup(sheet);
    return new ExcelPrintLayoutSnapshot(
        printLayout(sheet),
        new ExcelPrintSetupSnapshot(
            new ExcelPrintMarginsSnapshot(
                setup.margins().left(),
                setup.margins().right(),
                setup.margins().top(),
                setup.margins().bottom(),
                setup.margins().header(),
                setup.margins().footer()),
            setup.printGridlines(),
            setup.horizontallyCentered(),
            setup.verticallyCentered(),
            setup.paperSize(),
            setup.draft(),
            setup.blackAndWhite(),
            setup.copies(),
            setup.useFirstPageNumber(),
            setup.firstPageNumber(),
            setup.rowBreaks(),
            setup.columnBreaks()));
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
                pageSetup(sheetRef).orElse(null), "page setup must exist after fit-to-page toggle");
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

  private static void applySetup(XSSFSheet sheet, ExcelPrintSetup setup) {
    XSSFPrintSetup printSetup = sheet.getPrintSetup();
    applyMargins(sheet, setup.margins());
    sheet.setPrintGridlines(setup.printGridlines());
    sheet.setHorizontallyCenter(setup.horizontallyCentered());
    sheet.setVerticallyCenter(setup.verticallyCentered());
    printSetup.setPaperSize((short) setup.paperSize());
    printSetup.setDraft(setup.draft());
    printSetup.setNoColor(setup.blackAndWhite());
    printSetup.setCopies((short) setup.copies());
    printSetup.setUsePage(setup.useFirstPageNumber());
    printSetup.setPageStart((short) setup.firstPageNumber());
    replaceBreaks(
        sheet.getRowBreaks(), setup.rowBreaks(), sheet::removeRowBreak, sheet::setRowBreak);
    replaceBreaks(
        sheet.getColumnBreaks(),
        setup.columnBreaks(),
        sheet::removeColumnBreak,
        sheet::setColumnBreak);
  }

  private static void applyMargins(XSSFSheet sheet, ExcelPrintMargins margins) {
    sheet.setMargin(PageMargin.LEFT, margins.left());
    sheet.setMargin(PageMargin.RIGHT, margins.right());
    sheet.setMargin(PageMargin.TOP, margins.top());
    sheet.setMargin(PageMargin.BOTTOM, margins.bottom());
    sheet.setMargin(PageMargin.HEADER, margins.header());
    sheet.setMargin(PageMargin.FOOTER, margins.footer());
  }

  private static void replaceBreaks(
      int[] existingBreaks,
      java.util.List<Integer> authoredBreaks,
      java.util.function.IntConsumer remover,
      java.util.function.IntConsumer adder) {
    for (int existingBreak : existingBreaks) {
      remover.accept(existingBreak);
    }
    for (Integer authoredBreak : authoredBreaks) {
      adder.accept(authoredBreak);
    }
  }

  static Optional<String> storedPrintAreaFormula(XSSFSheet sheet) {
    return ExcelPrintLayoutReadSupport.storedPrintAreaFormula(sheet);
  }

  static Optional<String> rawPrintAreaFormula(XSSFWorkbook workbook, int sheetIndex) {
    return ExcelPrintLayoutReadSupport.rawPrintAreaFormula(workbook, sheetIndex);
  }

  static ExcelPrintLayout.Scaling scaling(XSSFSheet sheet) {
    return ExcelPrintLayoutReadSupport.scaling(sheet);
  }

  static void normalizeHeaderFooterNode(XSSFSheet sheet) {
    ExcelPrintLayoutNormalizationSupport.normalizeHeaderFooterNode(sheet);
  }

  static void normalizePageSetupNode(XSSFSheet sheet) {
    ExcelPrintLayoutNormalizationSupport.normalizePageSetupNode(sheet);
  }

  static boolean shouldUnsetPageSetupOrientation(XSSFSheet sheet, CTPageSetup pageSetup) {
    return ExcelPrintLayoutNormalizationSupport.shouldUnsetPageSetupOrientation(sheet, pageSetup);
  }

  static boolean isEmptyPageSetup(CTPageSetup pageSetup) {
    return ExcelPrintLayoutNormalizationSupport.isEmptyPageSetup(pageSetup);
  }

  static void normalizePageSetupProperties(XSSFSheet sheet) {
    ExcelPrintLayoutNormalizationSupport.normalizePageSetupProperties(sheet);
  }

  static void normalizePageMarginsNode(XSSFSheet sheet) {
    ExcelPrintLayoutNormalizationSupport.normalizePageMarginsNode(sheet);
  }

  static Optional<CTPageSetup> pageSetup(XSSFSheet sheet) {
    return ExcelPrintLayoutReadSupport.pageSetup(sheet);
  }

  static Optional<CTSheetPr> sheetPr(XSSFSheet sheet) {
    return ExcelPrintLayoutReadSupport.sheetPr(sheet);
  }

  static Optional<CTPageSetUpPr> pageSetUpPr(XSSFSheet sheet) {
    return ExcelPrintLayoutReadSupport.pageSetUpPr(sheet);
  }
}
