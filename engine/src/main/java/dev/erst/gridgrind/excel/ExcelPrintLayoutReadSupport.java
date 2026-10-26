package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.IntStream;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetUpPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation;

/** Reads supported print-layout state back out of the POI/OOXML sheet surface. */
final class ExcelPrintLayoutReadSupport {
  private ExcelPrintLayoutReadSupport() {}

  static PrintAreaSpec printArea(XSSFSheet sheet) {
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

  static Optional<String> storedPrintAreaFormula(XSSFSheet sheet) {
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    if (rawPrintAreaFormula(sheet.getWorkbook(), sheetIndex).isEmpty()) {
      return Optional.empty();
    }
    return nonBlank(sheet.getWorkbook().getPrintArea(sheetIndex));
  }

  static Optional<String> rawPrintAreaFormula(XSSFWorkbook workbook, int sheetIndex) {
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

  static ExcelPrintOrientation orientation(XSSFSheet sheet) {
    Optional<CTPageSetup> pageSetup = pageSetup(sheet);
    return pageSetup.isPresent()
            && pageSetup.orElseThrow().isSetOrientation()
            && pageSetup.orElseThrow().getOrientation() == STOrientation.LANDSCAPE
        ? ExcelPrintOrientation.LANDSCAPE
        : ExcelPrintOrientation.PORTRAIT;
  }

  static PrintScalingSpec scaling(XSSFSheet sheet) {
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

  static PrintTitleRowsSpec repeatingRows(XSSFSheet sheet) {
    CellRangeAddress repeatingRows = sheet.getRepeatingRows();
    if (repeatingRows == null) {
      return new PrintTitleRowsSpec.None();
    }
    return new PrintTitleRowsSpec.Band(repeatingRows.getFirstRow(), repeatingRows.getLastRow());
  }

  static PrintTitleColumnsSpec repeatingColumns(XSSFSheet sheet) {
    CellRangeAddress repeatingColumns = sheet.getRepeatingColumns();
    if (repeatingColumns == null) {
      return new PrintTitleColumnsSpec.None();
    }
    return new PrintTitleColumnsSpec.Band(
        repeatingColumns.getFirstColumn(), repeatingColumns.getLastColumn());
  }

  static HeaderFooterTextReport headerFooterText(HeaderFooter headerFooter) {
    return new HeaderFooterTextReport(
        Objects.toString(headerFooter.getLeft(), ""),
        Objects.toString(headerFooter.getCenter(), ""),
        Objects.toString(headerFooter.getRight(), ""));
  }

  static PrintSetupSpec setup(XSSFSheet sheet) {
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

  static Optional<CTPageSetup> pageSetup(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetPageSetup()
        ? Optional.of(sheet.getCTWorksheet().getPageSetup())
        : Optional.empty();
  }

  static Optional<CTSheetPr> sheetPr(XSSFSheet sheet) {
    return sheet.getCTWorksheet().isSetSheetPr()
        ? Optional.of(sheet.getCTWorksheet().getSheetPr())
        : Optional.empty();
  }

  static Optional<CTPageSetUpPr> pageSetUpPr(XSSFSheet sheet) {
    Optional<CTSheetPr> sheetPr = sheetPr(sheet);
    if (sheetPr.isEmpty() || !sheetPr.orElseThrow().isSetPageSetUpPr()) {
      return Optional.empty();
    }
    return Optional.of(sheetPr.orElseThrow().getPageSetUpPr());
  }

  private static Optional<String> nonBlank(String value) {
    return Optional.ofNullable(value).filter(candidate -> !candidate.isBlank());
  }
}
