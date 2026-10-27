package dev.erst.gridgrind.engine.print;

import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.engine.poi.CellBounds;
import java.util.Objects;
import java.util.Optional;
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

/** Owns print-layout mutation, factual readback and optional SpreadsheetML normalization. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class PrintAdapter {
  /** Applies the provided print layout as the authoritative supported print state. */
  public void setPrintLayout(XSSFSheet sheet, ResolvedPrintLayout layout) {
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
  public void clearPrintLayout(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    applyPrintArea(sheet, new PrintAreaSpec.None());
    applyRepeatingRows(sheet, new PrintTitleRowsSpec.None());
    applyRepeatingColumns(sheet, new PrintTitleColumnsSpec.None());
    applyHeader(sheet.getHeader(), HeaderFooterTextReport.blank());
    applyFooter(sheet.getFooter(), HeaderFooterTextReport.blank());
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
  public ResolvedPrintLayout printLayout(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return new ResolvedPrintLayout(
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
  public PrintLayoutReport printLayoutSnapshot(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    var layout = printLayout(sheet);
    PrintSetupSpec setup = layout.setup();
    return new PrintLayoutReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheet.getSheetName()),
        switch (layout.printArea()) {
          case PrintAreaSpec.None _ ->
              new dev.erst.gridgrind.core.model.print.PrintAreaReport.None();
          case PrintAreaSpec.Range value ->
              new dev.erst.gridgrind.core.model.print.PrintAreaReport.Range(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel(value.range().value()));
        },
        layout.orientation(),
        switch (layout.scaling()) {
          case PrintScalingSpec.Automatic _ ->
              new dev.erst.gridgrind.core.model.print.PrintScalingReport.Automatic();
          case PrintScalingSpec.Fit value ->
              new dev.erst.gridgrind.core.model.print.PrintScalingReport.Fit(
                  value.widthPages(), value.heightPages());
        },
        switch (layout.repeatingRows()) {
          case PrintTitleRowsSpec.None _ ->
              new dev.erst.gridgrind.core.model.print.PrintTitleRowsReport.None();
          case PrintTitleRowsSpec.Band value ->
              new dev.erst.gridgrind.core.model.print.PrintTitleRowsReport.Band(
                  value.firstRowIndex(), value.lastRowIndex());
        },
        switch (layout.repeatingColumns()) {
          case PrintTitleColumnsSpec.None _ ->
              new dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport.None();
          case PrintTitleColumnsSpec.Band value ->
              new dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport.Band(
                  value.firstColumnIndex(), value.lastColumnIndex());
        },
        layout.header(),
        layout.footer(),
        new PrintSetupReport(
            new PrintMarginsReport(
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

  private static void applyPrintArea(XSSFSheet sheet, PrintAreaSpec printArea) {
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    switch (printArea) {
      case PrintAreaSpec.None _ -> sheet.getWorkbook().removePrintArea(sheetIndex);
      case PrintAreaSpec.Range range -> {
        CellBounds parsed = CellBounds.parse(range.range().value());
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

  public static void applyScaling(XSSFSheet sheet, PrintScalingSpec scaling) {
    XSSFSheet sheetRef = Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFPrintSetup printSetup = sheetRef.getPrintSetup();
    switch (scaling) {
      case PrintScalingSpec.Automatic _ -> {
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
      case PrintScalingSpec.Fit fit -> {
        sheetRef.setFitToPage(true);
        printSetup.setFitWidth((short) fit.widthPages());
        printSetup.setFitHeight((short) fit.heightPages());
      }
    }
  }

  private static void applyRepeatingRows(XSSFSheet sheet, PrintTitleRowsSpec titleRows) {
    switch (titleRows) {
      case PrintTitleRowsSpec.None _ -> sheet.setRepeatingRows(null);
      case PrintTitleRowsSpec.Band band ->
          sheet.setRepeatingRows(
              new CellRangeAddress(band.firstRowIndex(), band.lastRowIndex(), -1, -1));
    }
  }

  private static void applyRepeatingColumns(XSSFSheet sheet, PrintTitleColumnsSpec titleColumns) {
    switch (titleColumns) {
      case PrintTitleColumnsSpec.None _ -> sheet.setRepeatingColumns(null);
      case PrintTitleColumnsSpec.Band band ->
          sheet.setRepeatingColumns(
              new CellRangeAddress(-1, -1, band.firstColumnIndex(), band.lastColumnIndex()));
    }
  }

  private static void applyHeader(Header header, HeaderFooterTextReport text) {
    header.setLeft(text.left());
    header.setCenter(text.center());
    header.setRight(text.right());
  }

  private static void applyFooter(Footer footer, HeaderFooterTextReport text) {
    footer.setLeft(text.left());
    footer.setCenter(text.center());
    footer.setRight(text.right());
  }

  private static void applySetup(XSSFSheet sheet, PrintSetupSpec setup) {
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

  private static void applyMargins(XSSFSheet sheet, PrintMarginsSpec margins) {
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

  public static PrintAreaSpec printArea(XSSFSheet sheet) {
    Optional<String> printArea = storedPrintAreaFormula(sheet);
    if (printArea.isEmpty()) {
      return new PrintAreaSpec.None();
    }
    AreaReference areaReference =
        new AreaReference(printArea.orElseThrow(), SpreadsheetVersion.EXCEL2007);
    return new PrintAreaSpec.Range(
        new RangeRef(
            new CellRangeAddress(
                    areaReference.getFirstCell().getRow(),
                    areaReference.getLastCell().getRow(),
                    areaReference.getFirstCell().getCol(),
                    areaReference.getLastCell().getCol())
                .formatAsString()));
  }

  public static Optional<String> storedPrintAreaFormula(XSSFSheet sheet) {
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    if (rawPrintAreaFormula(sheet.getWorkbook(), sheetIndex).isEmpty()) {
      return Optional.empty();
    }
    return nonBlank(sheet.getWorkbook().getPrintArea(sheetIndex));
  }

  public static Optional<String> rawPrintAreaFormula(XSSFWorkbook workbook, int sheetIndex) {
    if (!workbook.getCTWorkbook().isSetDefinedNames()) {
      return Optional.empty();
    }
    for (CTDefinedName definedName :
        workbook.getCTWorkbook().getDefinedNames().getDefinedNameList()) {
      if (XSSFName.BUILTIN_PRINT_AREA.equals(definedName.getName())
          && definedName.isSetLocalSheetId()
          && definedName.getLocalSheetId() == sheetIndex) {
        return nonBlank(definedName.getStringValue());
      }
    }
    return Optional.empty();
  }

  public static ExcelPrintOrientation orientation(XSSFSheet sheet) {
    Optional<CTPageSetup> pageSetup = pageSetup(sheet);
    return pageSetup.isPresent()
            && pageSetup.orElseThrow().isSetOrientation()
            && pageSetup.orElseThrow().getOrientation() == STOrientation.LANDSCAPE
        ? ExcelPrintOrientation.LANDSCAPE
        : ExcelPrintOrientation.PORTRAIT;
  }

  public static PrintScalingSpec scaling(XSSFSheet sheet) {
    Optional<CTPageSetUpPr> pageSetUpPr = pageSetUpPr(sheet);
    if (pageSetUpPr.isEmpty()
        || !pageSetUpPr.orElseThrow().isSetFitToPage()
        || !pageSetUpPr.orElseThrow().getFitToPage()) {
      return new PrintScalingSpec.Automatic();
    }
    Optional<CTPageSetup> pageSetup = pageSetup(sheet);
    return new PrintScalingSpec.Fit(
        pageSetup.isPresent() ? Math.toIntExact(pageSetup.orElseThrow().getFitToWidth()) : 1,
        pageSetup.isPresent() ? Math.toIntExact(pageSetup.orElseThrow().getFitToHeight()) : 1);
  }

  public static PrintTitleRowsSpec repeatingRows(XSSFSheet sheet) {
    CellRangeAddress repeatingRows = sheet.getRepeatingRows();
    if (repeatingRows == null) {
      return new PrintTitleRowsSpec.None();
    }
    return new PrintTitleRowsSpec.Band(repeatingRows.getFirstRow(), repeatingRows.getLastRow());
  }

  public static PrintTitleColumnsSpec repeatingColumns(XSSFSheet sheet) {
    CellRangeAddress repeatingColumns = sheet.getRepeatingColumns();
    if (repeatingColumns == null) {
      return new PrintTitleColumnsSpec.None();
    }
    return new PrintTitleColumnsSpec.Band(
        repeatingColumns.getFirstColumn(), repeatingColumns.getLastColumn());
  }

  public static HeaderFooterTextReport headerFooterText(HeaderFooter headerFooter) {
    return new HeaderFooterTextReport(
        Objects.toString(headerFooter.getLeft(), ""),
        Objects.toString(headerFooter.getCenter(), ""),
        Objects.toString(headerFooter.getRight(), ""));
  }

  public static PrintSetupSpec setup(XSSFSheet sheet) {
    Optional<CTPageSetup> pageSetup = pageSetup(sheet);
    PrintSetupSpec defaults = PrintSetupSpec.defaults();
    PrintMarginsSpec defaultMargins = defaults.margins();
    PrintMarginsSpec margins =
        sheet.getCTWorksheet().isSetPageMargins()
            ? new PrintMarginsSpec(
                sheet.getMargin(PageMargin.LEFT),
                sheet.getMargin(PageMargin.RIGHT),
                sheet.getMargin(PageMargin.TOP),
                sheet.getMargin(PageMargin.BOTTOM),
                sheet.getMargin(PageMargin.HEADER),
                sheet.getMargin(PageMargin.FOOTER))
            : defaultMargins;
    return new PrintSetupSpec(
        margins,
        sheet.isPrintGridlines(),
        sheet.getHorizontallyCenter(),
        sheet.getVerticallyCenter(),
        pageSetup.isPresent()
            ? Math.toIntExact(pageSetup.orElseThrow().getPaperSize())
            : defaults.paperSize(),
        pageSetup.isPresent() ? pageSetup.orElseThrow().getDraft() : defaults.draft(),
        pageSetup.isPresent()
            ? pageSetup.orElseThrow().getBlackAndWhite()
            : defaults.blackAndWhite(),
        pageSetup.isPresent()
            ? Math.toIntExact(pageSetup.orElseThrow().getCopies())
            : defaults.copies(),
        pageSetup.isPresent()
            ? pageSetup.orElseThrow().getUseFirstPageNumber()
            : defaults.useFirstPageNumber(),
        pageSetup.isPresent()
            ? Math.toIntExact(pageSetup.orElseThrow().getFirstPageNumber())
            : defaults.firstPageNumber(),
        IntStream.of(sheet.getRowBreaks()).boxed().toList(),
        IntStream.of(sheet.getColumnBreaks()).boxed().toList());
  }

  public static Optional<CTPageSetup> pageSetup(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetPageSetup()
        ? Optional.of(sheet.getCTWorksheet().getPageSetup())
        : Optional.empty();
  }

  public static Optional<CTSheetPr> sheetPr(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetSheetPr()
        ? Optional.of(sheet.getCTWorksheet().getSheetPr())
        : Optional.empty();
  }

  public static Optional<CTPageSetUpPr> pageSetUpPr(XSSFSheet sheet) {
    Optional<CTSheetPr> sheetPr = sheetPr(sheet);
    if (sheetPr.isEmpty() || !sheetPr.orElseThrow().isSetPageSetUpPr()) {
      return Optional.empty();
    }
    return Optional.of(sheetPr.orElseThrow().getPageSetUpPr());
  }

  private static Optional<String> nonBlank(String value) {
    return Optional.ofNullable(value).filter(candidate -> !candidate.isBlank());
  }

  public static void normalizePrintNodes(XSSFSheet sheet) {
    normalizeHeaderFooterNode(sheet);
    normalizePageSetupNode(sheet);
    normalizePageSetupProperties(sheet);
    normalizePageMarginsNode(sheet);
  }

  public static void normalizeHeaderFooterNode(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetHeaderFooter()) {
      return;
    }
    if (!headerFooterText(sheet.getHeader()).empty()) {
      return;
    }
    if (!headerFooterText(sheet.getFooter()).empty()) {
      return;
    }
    sheet.getCTWorksheet().unsetHeaderFooter();
  }

  public static void normalizePageSetupNode(XSSFSheet sheet) {
    Optional<CTPageSetup> pageSetup = pageSetup(sheet);
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

  public static boolean shouldUnsetPageSetupOrientation(XSSFSheet sheet, CTPageSetup pageSetup) {
    return pageSetup.isSetOrientation()
        && pageSetup.getOrientation() == STOrientation.PORTRAIT
        && !pageSetup.isSetFitToWidth()
        && !pageSetup.isSetFitToHeight();
  }

  public static boolean isEmptyPageSetup(CTPageSetup pageSetup) {
    PrintSetupSpec defaults = PrintSetupSpec.defaults();
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

  public static void normalizePageSetupProperties(XSSFSheet sheet) {
    Optional<CTSheetPr> sheetPr = sheetPr(sheet);
    if (sheetPr.isEmpty()) {
      return;
    }
    Optional<CTPageSetUpPr> pageSetUpPr = pageSetUpPr(sheet);
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

  public static void normalizePageMarginsNode(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetPageMargins()) {
      return;
    }
    PrintMarginsSpec defaults = PrintSetupSpec.defaults().margins();
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
