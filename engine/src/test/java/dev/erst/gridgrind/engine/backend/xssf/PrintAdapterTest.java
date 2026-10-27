package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.engine.print.PrintAdapter;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation;

/** Direct tests for the print-layout controller seam and its XML normalization behavior. */
class PrintAdapterTest {
  @Test
  void readsDefaultsFromFreshSheetsWithoutCreatingPrintNodes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      PrintAdapter controller = new PrintAdapter();

      assertTrue(sheet.getCTWorksheet().isSetPageMargins());
      assertFalse(sheet.getCTWorksheet().isSetPageSetup());
      assertEquals(factualDefaultPrintLayout(), controller.printLayout(sheet));
      assertFalse(sheet.getCTWorksheet().isSetPageSetup());
      assertTrue(sheet.getCTWorksheet().isSetPageMargins());
      controller.clearPrintLayout(sheet);

      assertEquals(factualDefaultPrintLayout(), controller.printLayout(sheet));
      assertFalse(sheet.getCTWorksheet().isSetPageSetup());
      assertFalse(sheet.getCTWorksheet().isSetPageMargins());
      assertNull(workbook.getPrintArea(workbook.getSheetIndex(sheet)));
    }
  }

  @Test
  void setAndClearPrintLayoutSynchronizesStateAndPrunesEmptyXmlNodes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      PrintAdapter controller = new PrintAdapter();
      ResolvedPrintLayout layout =
          new ResolvedPrintLayout(
              new PrintAreaSpec.Range(new RangeRef("A1:B20")),
              ExcelPrintOrientation.LANDSCAPE,
              new PrintScalingSpec.Fit(1, 0),
              new PrintTitleRowsSpec.Band(0, 0),
              new PrintTitleColumnsSpec.Band(0, 1),
              new HeaderFooterTextReport("Budget", "", ""),
              new HeaderFooterTextReport("", "Page &P", ""));

      controller.setPrintLayout(sheet, layout);

      assertEquals(layout, controller.printLayout(sheet));
      assertTrue(sheet.getCTWorksheet().isSetHeaderFooter());
      assertTrue(sheet.getCTWorksheet().isSetPageSetup());
      assertTrue(sheet.getCTWorksheet().isSetSheetPr());

      controller.clearPrintLayout(sheet);

      assertNull(workbook.getPrintArea(workbook.getSheetIndex(sheet)));
      assertNull(sheet.getRepeatingRows());
      assertNull(sheet.getRepeatingColumns());
    }
  }

  @Test
  void setAndClearPrintGridlinesRoundTripsThroughPrintSetup() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      PrintAdapter controller = new PrintAdapter();
      ResolvedPrintLayout layout =
          new ResolvedPrintLayout(
              new PrintAreaSpec.None(),
              ExcelPrintOrientation.PORTRAIT,
              new PrintScalingSpec.Automatic(),
              new PrintTitleRowsSpec.None(),
              new PrintTitleColumnsSpec.None(),
              HeaderFooterTextReport.blank(),
              HeaderFooterTextReport.blank(),
              new PrintSetupSpec(
                  new PrintMarginsSpec(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
                  true,
                  false,
                  false,
                  1,
                  false,
                  false,
                  1,
                  false,
                  1,
                  java.util.List.of(),
                  java.util.List.of()));

      controller.setPrintLayout(sheet, layout);

      assertTrue(sheet.isPrintGridlines());
      assertTrue(controller.printLayout(sheet).setup().printGridlines());

      controller.clearPrintLayout(sheet);

      assertFalse(sheet.isPrintGridlines());
      assertFalse(controller.printLayout(sheet).setup().printGridlines());
    }
  }

  @Test
  void preservesSupportedPageSetupAttributesWhenOrientationAndScalingAreDefault() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      PrintAdapter controller = new PrintAdapter();
      ResolvedPrintLayout layout =
          new ResolvedPrintLayout(
              new PrintAreaSpec.None(),
              ExcelPrintOrientation.PORTRAIT,
              new PrintScalingSpec.Automatic(),
              new PrintTitleRowsSpec.None(),
              new PrintTitleColumnsSpec.None(),
              HeaderFooterTextReport.blank(),
              HeaderFooterTextReport.blank(),
              new PrintSetupSpec(
                  new PrintMarginsSpec(0.35d, 0.55d, 0.6d, 0.45d, 0.3d, 0.3d),
                  false,
                  true,
                  true,
                  8,
                  true,
                  true,
                  2,
                  true,
                  4,
                  java.util.List.of(6),
                  java.util.List.of(3)));

      controller.setPrintLayout(sheet, layout);

      assertTrue(sheet.getCTWorksheet().isSetPageSetup());
      assertEquals(layout, controller.printLayout(sheet));
      assertFalse(PrintAdapter.isEmptyPageSetup(sheet.getCTWorksheet().getPageSetup()));

      PrintAdapter.normalizePageSetupNode(sheet);

      assertTrue(sheet.getCTWorksheet().isSetPageSetup());
      assertEquals(layout.setup(), controller.printLayout(sheet).setup());
    }
  }

  @Test
  void normalizesBlankAndNegativeStoredStateWhilePreservingPrinterDefaults() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      XSSFSheet rowsOnlySheet = workbook.createSheet("RowsOnly");
      PrintAdapter controller = new PrintAdapter();

      var definedNames =
          workbook.getCTWorkbook().isSetDefinedNames()
              ? workbook.getCTWorkbook().getDefinedNames()
              : workbook.getCTWorkbook().addNewDefinedNames();
      var definedName = definedNames.addNewDefinedName();
      definedName.setName(XSSFName.BUILTIN_PRINT_AREA);
      definedName.setLocalSheetId(workbook.getSheetIndex(sheet));
      definedName.setStringValue(" ");
      sheet.setRepeatingColumns(new CellRangeAddress(-1, -1, 0, 1));
      rowsOnlySheet.setRepeatingRows(new CellRangeAddress(0, 1, -1, -1));
      sheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr().setFitToPage(false);
      sheet.getCTWorksheet().addNewPageSetup().setUsePrinterDefaults(true);

      ResolvedPrintLayout layout = controller.printLayout(sheet);
      ResolvedPrintLayout rowsOnlyLayout = controller.printLayout(rowsOnlySheet);

      assertEquals(new PrintAreaSpec.None(), layout.printArea());
      assertEquals(new PrintScalingSpec.Automatic(), layout.scaling());
      assertEquals(new PrintTitleRowsSpec.None(), layout.repeatingRows());
      assertEquals(new PrintTitleColumnsSpec.Band(0, 1), layout.repeatingColumns());
      assertEquals(new PrintTitleRowsSpec.Band(0, 1), rowsOnlyLayout.repeatingRows());
      assertEquals(new PrintTitleColumnsSpec.None(), rowsOnlyLayout.repeatingColumns());

      controller.clearPrintLayout(sheet);
    }
  }

  @Test
  void exposesInternalPrintAreaHelpersForBlankAndForeignDefinitions() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet budget = workbook.createSheet("Budget");
      XSSFSheet other = workbook.createSheet("Other");

      assertTrue(PrintAdapter.pageSetup(budget).isEmpty());
      assertTrue(PrintAdapter.sheetPr(budget).isEmpty());
      assertTrue(PrintAdapter.pageSetUpPr(budget).isEmpty());

      var definedNames = workbook.getCTWorkbook().addNewDefinedNames();
      var unscopedPrintArea = definedNames.addNewDefinedName();
      unscopedPrintArea.setName(XSSFName.BUILTIN_PRINT_AREA);
      unscopedPrintArea.setStringValue("Budget!$C$1:$D$2");
      var otherPrintArea = definedNames.addNewDefinedName();
      otherPrintArea.setName(XSSFName.BUILTIN_PRINT_AREA);
      otherPrintArea.setLocalSheetId(workbook.getSheetIndex(other));
      otherPrintArea.setStringValue("Other!$A$1:$B$2");

      assertTrue(PrintAdapter.storedPrintAreaFormula(budget).isEmpty());

      var blankBudgetPrintArea = definedNames.addNewDefinedName();
      blankBudgetPrintArea.setName(XSSFName.BUILTIN_PRINT_AREA);
      blankBudgetPrintArea.setLocalSheetId(workbook.getSheetIndex(budget));
      blankBudgetPrintArea.setStringValue(" ");

      assertTrue(PrintAdapter.storedPrintAreaFormula(budget).isEmpty());
    }
  }

  @Test
  void exposesInternalScalingHelpersForAutomaticEdgeStates() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet budget = workbook.createSheet("Budget");
      var scalingSheet = workbook.createSheet("Scaling");

      PrintAdapter.applyScaling(budget, new PrintScalingSpec.Automatic());
      scalingSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr().setFitToPage(true);
      scalingSheet.getCTWorksheet().addNewPageSetup().setFitToWidth((short) 2);
      scalingSheet.getCTWorksheet().getPageSetup().setFitToHeight((short) 3);
      PrintAdapter.applyScaling(scalingSheet, new PrintScalingSpec.Automatic());
      assertFalse(PrintAdapter.pageSetup(scalingSheet).orElseThrow().isSetFitToWidth());
      assertFalse(PrintAdapter.pageSetup(scalingSheet).orElseThrow().isSetFitToHeight());

      var automaticSheet = workbook.createSheet("Automatic");
      assertEquals(new PrintScalingSpec.Automatic(), PrintAdapter.scaling(automaticSheet));
      automaticSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr();
      assertEquals(new PrintScalingSpec.Automatic(), PrintAdapter.scaling(automaticSheet));

      var fitWithoutPageSetup = workbook.createSheet("FitWithoutPageSetup");
      fitWithoutPageSetup.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr().setFitToPage(true);
      assertEquals(new PrintScalingSpec.Fit(1, 1), PrintAdapter.scaling(fitWithoutPageSetup));

      var fitSheet = workbook.createSheet("Fit");
      PrintAdapter.applyScaling(fitSheet, new PrintScalingSpec.Fit(2, 3));
      assertTrue(fitSheet.getCTWorksheet().isSetPageSetup());
      assertEquals(new PrintScalingSpec.Fit(2, 3), PrintAdapter.scaling(fitSheet));
    }
  }

  @Test
  void setPrintLayoutReplacesExistingRowAndColumnBreaks() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      PrintAdapter controller = new PrintAdapter();
      sheet.setRowBreak(1);
      sheet.setRowBreak(3);
      sheet.setColumnBreak(2);

      controller.setPrintLayout(
          sheet,
          new ResolvedPrintLayout(
              new PrintAreaSpec.None(),
              ExcelPrintOrientation.PORTRAIT,
              new PrintScalingSpec.Automatic(),
              new PrintTitleRowsSpec.None(),
              new PrintTitleColumnsSpec.None(),
              HeaderFooterTextReport.blank(),
              HeaderFooterTextReport.blank(),
              new PrintSetupSpec(
                  PrintSetupSpec.defaults().margins(),
                  false,
                  false,
                  false,
                  1,
                  false,
                  false,
                  1,
                  false,
                  1,
                  java.util.List.of(4, 6),
                  java.util.List.of(1, 5))));

      assertEquals(
          java.util.List.of(4, 6), java.util.Arrays.stream(sheet.getRowBreaks()).boxed().toList());
      assertEquals(
          java.util.List.of(1, 5),
          java.util.Arrays.stream(sheet.getColumnBreaks()).boxed().toList());
    }
  }

  @Test
  void exposesInternalHeaderFooterHelpersForBlankAndNonBlankNodes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet other = workbook.createSheet("Other");
      var blankHeaderFooterSheet = workbook.createSheet("HeaderFooter");
      PrintAdapter.normalizeHeaderFooterNode(other);
      blankHeaderFooterSheet.getHeader().setLeft("");
      blankHeaderFooterSheet.getFooter().setRight("");
      assertTrue(blankHeaderFooterSheet.getCTWorksheet().isSetHeaderFooter());
      PrintAdapter.normalizeHeaderFooterNode(blankHeaderFooterSheet);
      assertFalse(blankHeaderFooterSheet.getCTWorksheet().isSetHeaderFooter());
      var nonBlankFooterSheet = workbook.createSheet("NonBlankFooter");
      nonBlankFooterSheet.getFooter().setRight("keep");
      PrintAdapter.normalizeHeaderFooterNode(nonBlankFooterSheet);
      assertTrue(nonBlankFooterSheet.getCTWorksheet().isSetHeaderFooter());
    }
  }

  @Test
  void exposesInternalPageSetupHelpersForNormalizationEdgeStates() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertPageSetupOrientationNormalization(workbook);
      assertEmptyAndNonEmptyPageSetupStates(workbook);
      assertPageSetupPropertiesNormalization(workbook);
      assertPageMarginsNormalization(workbook);
    }
  }

  private static void assertPageSetupOrientationNormalization(XSSFWorkbook workbook) {
    var pageSetupSheet = workbook.createSheet("PageSetup");
    var pageSetup = pageSetupSheet.getCTWorksheet().addNewPageSetup();
    pageSetup.setOrientation(STOrientation.PORTRAIT);
    assertTrue(PrintAdapter.shouldUnsetPageSetupOrientation(pageSetupSheet, pageSetup));
    PrintAdapter.normalizePageSetupNode(pageSetupSheet);
    assertFalse(pageSetupSheet.getCTWorksheet().isSetPageSetup());

    var noPageSetupSheet = workbook.createSheet("NoPageSetup");
    PrintAdapter.normalizePageSetupNode(noPageSetupSheet);
    assertTrue(PrintAdapter.pageSetup(noPageSetupSheet).isEmpty());

    var fitWidthSheet = workbook.createSheet("FitWidth");
    var fitWidthPageSetup = fitWidthSheet.getCTWorksheet().addNewPageSetup();
    fitWidthPageSetup.setOrientation(STOrientation.PORTRAIT);
    fitWidthPageSetup.setFitToWidth((short) 1);
    assertFalse(PrintAdapter.shouldUnsetPageSetupOrientation(fitWidthSheet, fitWidthPageSetup));

    var fitHeightSheet = workbook.createSheet("FitHeight");
    var fitHeightPageSetup = fitHeightSheet.getCTWorksheet().addNewPageSetup();
    fitHeightPageSetup.setOrientation(STOrientation.PORTRAIT);
    fitHeightPageSetup.setFitToHeight((short) 1);
    assertFalse(PrintAdapter.shouldUnsetPageSetupOrientation(fitHeightSheet, fitHeightPageSetup));

    var noOrientationSheet = workbook.createSheet("NoOrientation");
    var noOrientationPageSetup = noOrientationSheet.getCTWorksheet().addNewPageSetup();
    assertEquals(
        ExcelPrintOrientation.PORTRAIT,
        new PrintAdapter().printLayout(noOrientationSheet).orientation());
    assertFalse(
        PrintAdapter.shouldUnsetPageSetupOrientation(noOrientationSheet, noOrientationPageSetup));

    var portraitSheet = workbook.createSheet("Portrait");
    portraitSheet.getCTWorksheet().addNewPageSetup().setOrientation(STOrientation.PORTRAIT);
    assertEquals(
        ExcelPrintOrientation.PORTRAIT,
        new PrintAdapter().printLayout(portraitSheet).orientation());

    var landscapeSheet = workbook.createSheet("Landscape");
    var landscapePageSetup = landscapeSheet.getCTWorksheet().addNewPageSetup();
    landscapePageSetup.setOrientation(STOrientation.LANDSCAPE);
    assertFalse(PrintAdapter.shouldUnsetPageSetupOrientation(landscapeSheet, landscapePageSetup));
  }

  private static void assertEmptyAndNonEmptyPageSetupStates(XSSFWorkbook workbook) {
    assertFitAndPrinterDefaultPageSetupStates(workbook);
    assertPaperAndBooleanPageSetupStates(workbook);
    assertCopiesAndFirstPageNumberPageSetupStates(workbook);
  }

  private static void assertFitAndPrinterDefaultPageSetupStates(XSSFWorkbook workbook) {
    var fitWidthOnlySheet = workbook.createSheet("FitWidthOnly");
    var fitWidthOnlyPageSetup = fitWidthOnlySheet.getCTWorksheet().addNewPageSetup();
    fitWidthOnlyPageSetup.setFitToWidth((short) 1);
    assertFalse(PrintAdapter.isEmptyPageSetup(fitWidthOnlyPageSetup));

    var fitHeightOnlySheet = workbook.createSheet("FitHeightOnly");
    var fitHeightOnlyPageSetup = fitHeightOnlySheet.getCTWorksheet().addNewPageSetup();
    fitHeightOnlyPageSetup.setFitToHeight((short) 1);
    assertFalse(PrintAdapter.isEmptyPageSetup(fitHeightOnlyPageSetup));

    var printerDefaultsSheet = workbook.createSheet("PrinterDefaults");
    var printerDefaults = printerDefaultsSheet.getCTWorksheet().addNewPageSetup();
    printerDefaults.setUsePrinterDefaults(true);
    assertFalse(PrintAdapter.isEmptyPageSetup(printerDefaults));

    var advancedSetupSheet = workbook.createSheet("AdvancedSetup");
    var advancedSetup = advancedSetupSheet.getCTWorksheet().addNewPageSetup();
    advancedSetup.setPaperSize((short) 8);
    advancedSetup.setDraft(true);
    advancedSetup.setBlackAndWhite(true);
    advancedSetup.setCopies((short) 2);
    advancedSetup.setUseFirstPageNumber(true);
    advancedSetup.setFirstPageNumber((short) 4);
    assertFalse(PrintAdapter.isEmptyPageSetup(advancedSetup));
  }

  private static void assertPaperAndBooleanPageSetupStates(XSSFWorkbook workbook) {
    var paperSizeDefaultSheet = workbook.createSheet("PaperSizeDefault");
    var paperSizeDefault = paperSizeDefaultSheet.getCTWorksheet().addNewPageSetup();
    paperSizeDefault.setPaperSize((short) 1);
    assertTrue(PrintAdapter.isEmptyPageSetup(paperSizeDefault));

    var draftTrueSheet = workbook.createSheet("DraftTrue");
    var draftTrue = draftTrueSheet.getCTWorksheet().addNewPageSetup();
    draftTrue.setDraft(true);
    assertFalse(PrintAdapter.isEmptyPageSetup(draftTrue));

    var draftDefaultSheet = workbook.createSheet("DraftDefault");
    var draftDefault = draftDefaultSheet.getCTWorksheet().addNewPageSetup();
    draftDefault.setDraft(false);
    assertTrue(PrintAdapter.isEmptyPageSetup(draftDefault));

    var bwTrueSheet = workbook.createSheet("BwTrue");
    var bwTrue = bwTrueSheet.getCTWorksheet().addNewPageSetup();
    bwTrue.setBlackAndWhite(true);
    assertFalse(PrintAdapter.isEmptyPageSetup(bwTrue));

    var bwDefaultSheet = workbook.createSheet("BwDefault");
    var bwDefault = bwDefaultSheet.getCTWorksheet().addNewPageSetup();
    bwDefault.setBlackAndWhite(false);
    assertTrue(PrintAdapter.isEmptyPageSetup(bwDefault));
  }

  private static void assertCopiesAndFirstPageNumberPageSetupStates(XSSFWorkbook workbook) {
    var copiesTwoSheet = workbook.createSheet("CopiesTwo");
    var copiesTwo = copiesTwoSheet.getCTWorksheet().addNewPageSetup();
    copiesTwo.setCopies((short) 2);
    assertFalse(PrintAdapter.isEmptyPageSetup(copiesTwo));

    var copiesDefaultSheet = workbook.createSheet("CopiesDefault");
    var copiesDefault = copiesDefaultSheet.getCTWorksheet().addNewPageSetup();
    copiesDefault.setCopies((short) 1);
    assertTrue(PrintAdapter.isEmptyPageSetup(copiesDefault));

    var useFirstPageTrueSheet = workbook.createSheet("UseFirstPageTrue");
    var useFirstPageTrue = useFirstPageTrueSheet.getCTWorksheet().addNewPageSetup();
    useFirstPageTrue.setUseFirstPageNumber(true);
    assertFalse(PrintAdapter.isEmptyPageSetup(useFirstPageTrue));

    var useFirstPageDefaultSheet = workbook.createSheet("UseFirstPageDefault");
    var useFirstPageDefault = useFirstPageDefaultSheet.getCTWorksheet().addNewPageSetup();
    useFirstPageDefault.setUseFirstPageNumber(false);
    assertTrue(PrintAdapter.isEmptyPageSetup(useFirstPageDefault));

    var firstPageNumTwoSheet = workbook.createSheet("FirstPageNumTwo");
    var firstPageNumTwo = firstPageNumTwoSheet.getCTWorksheet().addNewPageSetup();
    firstPageNumTwo.setFirstPageNumber((short) 2);
    assertFalse(PrintAdapter.isEmptyPageSetup(firstPageNumTwo));

    var firstPageNumDefaultSheet = workbook.createSheet("FirstPageNumDefault");
    var firstPageNumDefault = firstPageNumDefaultSheet.getCTWorksheet().addNewPageSetup();
    firstPageNumDefault.setFirstPageNumber((short) 1);
    assertTrue(PrintAdapter.isEmptyPageSetup(firstPageNumDefault));
  }

  private static void assertPageSetupPropertiesNormalization(XSSFWorkbook workbook) {
    var missingSheetPr = workbook.createSheet("MissingSheetPr");
    PrintAdapter.normalizePageSetupProperties(missingSheetPr);
    assertTrue(PrintAdapter.sheetPr(missingSheetPr).isEmpty());

    var staleSheetPr = workbook.createSheet("StaleSheetPr");
    staleSheetPr.getCTWorksheet().addNewSheetPr();
    PrintAdapter.normalizePageSetupProperties(staleSheetPr);
    assertFalse(staleSheetPr.getCTWorksheet().isSetSheetPr());
    assertTrue(PrintAdapter.pageSetUpPr(staleSheetPr).isEmpty());

    var unsetFitSheet = workbook.createSheet("UnsetFit");
    unsetFitSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr();
    PrintAdapter.normalizePageSetupProperties(unsetFitSheet);
    assertFalse(unsetFitSheet.getCTWorksheet().isSetSheetPr());

    var falseFitSheet = workbook.createSheet("FalseFit");
    falseFitSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr().setFitToPage(false);
    PrintAdapter.normalizePageSetupProperties(falseFitSheet);
    assertFalse(falseFitSheet.getCTWorksheet().isSetSheetPr());
  }

  private static void assertPageMarginsNormalization(XSSFWorkbook workbook) {
    PrintMarginsSpec defaults = PrintSetupSpec.defaults().margins();

    var defaultMarginsSheet = workbook.createSheet("DefaultMargins");
    PrintAdapter.normalizePageMarginsNode(defaultMarginsSheet);
    assertFalse(defaultMarginsSheet.getCTWorksheet().isSetPageMargins());

    var leftChangedSheet = workbook.createSheet("LeftMarginChanged");
    leftChangedSheet.setMargin(PageMargin.LEFT, defaults.left() + 0.1d);
    PrintAdapter.normalizePageMarginsNode(leftChangedSheet);
    assertTrue(leftChangedSheet.getCTWorksheet().isSetPageMargins());

    var rightChangedSheet = workbook.createSheet("RightMarginChanged");
    rightChangedSheet.setMargin(PageMargin.RIGHT, defaults.right() + 0.1d);
    PrintAdapter.normalizePageMarginsNode(rightChangedSheet);
    assertTrue(rightChangedSheet.getCTWorksheet().isSetPageMargins());

    var topChangedSheet = workbook.createSheet("TopMarginChanged");
    topChangedSheet.setMargin(PageMargin.TOP, defaults.top() + 0.1d);
    PrintAdapter.normalizePageMarginsNode(topChangedSheet);
    assertTrue(topChangedSheet.getCTWorksheet().isSetPageMargins());

    var bottomChangedSheet = workbook.createSheet("BottomMarginChanged");
    bottomChangedSheet.setMargin(PageMargin.BOTTOM, defaults.bottom() + 0.1d);
    PrintAdapter.normalizePageMarginsNode(bottomChangedSheet);
    assertTrue(bottomChangedSheet.getCTWorksheet().isSetPageMargins());

    var headerChangedSheet = workbook.createSheet("HeaderMarginChanged");
    headerChangedSheet.setMargin(PageMargin.HEADER, defaults.header() + 0.1d);
    PrintAdapter.normalizePageMarginsNode(headerChangedSheet);
    assertTrue(headerChangedSheet.getCTWorksheet().isSetPageMargins());

    var footerChangedSheet = workbook.createSheet("FooterMarginChanged");
    footerChangedSheet.setMargin(PageMargin.FOOTER, defaults.footer() + 0.1d);
    PrintAdapter.normalizePageMarginsNode(footerChangedSheet);
    assertTrue(footerChangedSheet.getCTWorksheet().isSetPageMargins());
  }

  private static ResolvedPrintLayout factualDefaultPrintLayout() {
    return new ResolvedPrintLayout(
        new PrintAreaSpec.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingSpec.Automatic(),
        new PrintTitleRowsSpec.None(),
        new PrintTitleColumnsSpec.None(),
        HeaderFooterTextReport.blank(),
        HeaderFooterTextReport.blank(),
        new PrintSetupSpec(
            new PrintMarginsSpec(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
            false,
            false,
            false,
            1,
            false,
            false,
            1,
            false,
            1,
            java.util.List.of(),
            java.util.List.of()));
  }
}
