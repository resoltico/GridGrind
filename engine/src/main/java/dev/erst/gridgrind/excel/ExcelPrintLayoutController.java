package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.util.Objects;
import java.util.stream.IntStream;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.PageMargin;
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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation;

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
    normalizePrintNodes(sheet);
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
    normalizePrintNodes(sheet);
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
        headerFooterText(sheet.getFooter()),
        setup(sheet));
  }

  /** Returns the full factual print-layout snapshot currently stored for the provided sheet. */
  ExcelPrintLayoutSnapshot printLayoutSnapshot(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelPrintSetup setup = setup(sheet);
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
      return java.util.Optional.<String>empty().orElse(null);
    }
    for (CTDefinedName definedName :
        workbook.getCTWorkbook().getDefinedNames().getDefinedNameList()) {
      if (XSSFName.BUILTIN_PRINT_AREA.equals(definedName.getName())
          && definedName.isSetLocalSheetId()
          && definedName.getLocalSheetId() == sheetIndex) {
        return definedName.getStringValue();
      }
    }
    return java.util.Optional.<String>empty().orElse(null);
  }

  private static ExcelPrintOrientation orientation(XSSFSheet sheet) {
    CTPageSetup pageSetup = pageSetupOrNull(sheet);
    return pageSetup != null
            && pageSetup.isSetOrientation()
            && pageSetup.getOrientation() == STOrientation.LANDSCAPE
        ? ExcelPrintOrientation.LANDSCAPE
        : ExcelPrintOrientation.PORTRAIT;
  }

  static ExcelPrintLayout.Scaling scaling(XSSFSheet sheet) {
    CTPageSetUpPr pageSetUpPr = pageSetUpPrOrNull(sheet);
    if (pageSetUpPr == null || !pageSetUpPr.isSetFitToPage() || !pageSetUpPr.getFitToPage()) {
      return new ExcelPrintLayout.Scaling.Automatic();
    }
    CTPageSetup pageSetup = pageSetupOrNull(sheet);
    return new ExcelPrintLayout.Scaling.Fit(
        pageSetup != null ? Math.toIntExact(pageSetup.getFitToWidth()) : 1,
        pageSetup != null ? Math.toIntExact(pageSetup.getFitToHeight()) : 1);
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

  private static ExcelPrintSetup setup(XSSFSheet sheet) {
    CTPageSetup pageSetup = pageSetupOrNull(sheet);
    ExcelPrintSetup defaults = ExcelPrintSetup.defaults();
    ExcelPrintMargins defaultMargins = defaults.margins();
    ExcelPrintMargins margins =
        sheet.getCTWorksheet().isSetPageMargins()
            ? new ExcelPrintMargins(
                sheet.getMargin(PageMargin.LEFT),
                sheet.getMargin(PageMargin.RIGHT),
                sheet.getMargin(PageMargin.TOP),
                sheet.getMargin(PageMargin.BOTTOM),
                sheet.getMargin(PageMargin.HEADER),
                sheet.getMargin(PageMargin.FOOTER))
            : defaultMargins;
    return new ExcelPrintSetup(
        margins,
        sheet.isPrintGridlines(),
        sheet.getHorizontallyCenter(),
        sheet.getVerticallyCenter(),
        pageSetup != null ? Math.toIntExact(pageSetup.getPaperSize()) : defaults.paperSize(),
        pageSetup != null ? pageSetup.getDraft() : defaults.draft(),
        pageSetup != null ? pageSetup.getBlackAndWhite() : defaults.blackAndWhite(),
        pageSetup != null ? Math.toIntExact(pageSetup.getCopies()) : defaults.copies(),
        pageSetup != null ? pageSetup.getUseFirstPageNumber() : defaults.useFirstPageNumber(),
        pageSetup != null
            ? Math.toIntExact(pageSetup.getFirstPageNumber())
            : defaults.firstPageNumber(),
        IntStream.of(sheet.getRowBreaks()).boxed().toList(),
        IntStream.of(sheet.getColumnBreaks()).boxed().toList());
  }

  private static void normalizePrintNodes(XSSFSheet sheet) {
    normalizeHeaderFooterNode(sheet);
    normalizePageSetupNode(sheet);
    normalizePageSetupProperties(sheet);
    normalizePageMarginsNode(sheet);
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

  static CTPageSetup pageSetupOrNull(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetPageSetup() ? sheet.getCTWorksheet().getPageSetup() : null;
  }

  static CTSheetPr sheetPrOrNull(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetSheetPr() ? sheet.getCTWorksheet().getSheetPr() : null;
  }

  static CTPageSetUpPr pageSetUpPrOrNull(XSSFSheet sheet) {
    CTSheetPr sheetPr = sheetPrOrNull(sheet);
    if (sheetPr == null || !sheetPr.isSetPageSetUpPr()) {
      return java.util.Optional.<CTPageSetUpPr>empty().orElse(null);
    }
    return sheetPr.getPageSetUpPr();
  }
}
