package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation;

/** Direct tests for the print-layout controller seam and its XML normalization behavior. */
class ExcelPrintLayoutControllerTest {
  @Test
  void readsDefaultsFromFreshSheetsWithoutCreatingPrintNodes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      ExcelPrintLayoutController controller = new ExcelPrintLayoutController();

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
      ExcelPrintLayoutController controller = new ExcelPrintLayoutController();
      ExcelPrintLayout layout =
          new ExcelPrintLayout(
              new ExcelPrintLayout.Area.Range("A1:B20"),
              ExcelPrintOrientation.LANDSCAPE,
              new ExcelPrintLayout.Scaling.Fit(1, 0),
              new ExcelPrintLayout.TitleRows.Band(0, 0),
              new ExcelPrintLayout.TitleColumns.Band(0, 1),
              new ExcelHeaderFooterText("Budget", "", ""),
              new ExcelHeaderFooterText("", "Page &P", ""));

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
      ExcelPrintLayoutController controller = new ExcelPrintLayoutController();
      ExcelPrintLayout layout =
          new ExcelPrintLayout(
              new ExcelPrintLayout.Area.None(),
              ExcelPrintOrientation.PORTRAIT,
              new ExcelPrintLayout.Scaling.Automatic(),
              new ExcelPrintLayout.TitleRows.None(),
              new ExcelPrintLayout.TitleColumns.None(),
              ExcelHeaderFooterText.blank(),
              ExcelHeaderFooterText.blank(),
              new ExcelPrintSetup(
                  new ExcelPrintMargins(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
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
      ExcelPrintLayoutController controller = new ExcelPrintLayoutController();
      ExcelPrintLayout layout =
          new ExcelPrintLayout(
              new ExcelPrintLayout.Area.None(),
              ExcelPrintOrientation.PORTRAIT,
              new ExcelPrintLayout.Scaling.Automatic(),
              new ExcelPrintLayout.TitleRows.None(),
              new ExcelPrintLayout.TitleColumns.None(),
              ExcelHeaderFooterText.blank(),
              ExcelHeaderFooterText.blank(),
              new ExcelPrintSetup(
                  new ExcelPrintMargins(0.35d, 0.55d, 0.6d, 0.45d, 0.3d, 0.3d),
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
      assertFalse(
          ExcelPrintLayoutController.isEmptyPageSetup(sheet.getCTWorksheet().getPageSetup()));

      ExcelPrintLayoutController.normalizePageSetupNode(sheet);

      assertTrue(sheet.getCTWorksheet().isSetPageSetup());
      assertEquals(layout.setup(), controller.printLayout(sheet).setup());
    }
  }

  @Test
  void normalizesBlankAndNegativeStoredStateWhilePreservingPrinterDefaults() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      XSSFSheet rowsOnlySheet = workbook.createSheet("RowsOnly");
      ExcelPrintLayoutController controller = new ExcelPrintLayoutController();

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

      ExcelPrintLayout layout = controller.printLayout(sheet);
      ExcelPrintLayout rowsOnlyLayout = controller.printLayout(rowsOnlySheet);

      assertEquals(new ExcelPrintLayout.Area.None(), layout.printArea());
      assertEquals(new ExcelPrintLayout.Scaling.Automatic(), layout.scaling());
      assertEquals(new ExcelPrintLayout.TitleRows.None(), layout.repeatingRows());
      assertEquals(new ExcelPrintLayout.TitleColumns.Band(0, 1), layout.repeatingColumns());
      assertEquals(new ExcelPrintLayout.TitleRows.Band(0, 1), rowsOnlyLayout.repeatingRows());
      assertEquals(new ExcelPrintLayout.TitleColumns.None(), rowsOnlyLayout.repeatingColumns());

      controller.clearPrintLayout(sheet);
    }
  }

  @Test
  void exposesInternalPrintAreaHelpersForBlankAndForeignDefinitions() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet budget = workbook.createSheet("Budget");
      XSSFSheet other = workbook.createSheet("Other");

      assertNull(ExcelPrintLayoutController.pageSetupOrNull(budget));
      assertNull(ExcelPrintLayoutController.sheetPrOrNull(budget));
      assertNull(ExcelPrintLayoutController.pageSetUpPrOrNull(budget));

      var definedNames = workbook.getCTWorkbook().addNewDefinedNames();
      var unscopedPrintArea = definedNames.addNewDefinedName();
      unscopedPrintArea.setName(XSSFName.BUILTIN_PRINT_AREA);
      unscopedPrintArea.setStringValue("Budget!$C$1:$D$2");
      var otherPrintArea = definedNames.addNewDefinedName();
      otherPrintArea.setName(XSSFName.BUILTIN_PRINT_AREA);
      otherPrintArea.setLocalSheetId(workbook.getSheetIndex(other));
      otherPrintArea.setStringValue("Other!$A$1:$B$2");

      assertNull(ExcelPrintLayoutController.storedPrintAreaFormulaOrNull(budget));

      var blankBudgetPrintArea = definedNames.addNewDefinedName();
      blankBudgetPrintArea.setName(XSSFName.BUILTIN_PRINT_AREA);
      blankBudgetPrintArea.setLocalSheetId(workbook.getSheetIndex(budget));
      blankBudgetPrintArea.setStringValue(" ");

      assertEquals(" ", ExcelPrintLayoutController.storedPrintAreaFormulaOrNull(budget));
    }
  }

  @Test
  void exposesInternalScalingHelpersForAutomaticEdgeStates() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet budget = workbook.createSheet("Budget");
      var scalingSheet = workbook.createSheet("Scaling");

      ExcelPrintLayoutController.applyScaling(budget, new ExcelPrintLayout.Scaling.Automatic());
      scalingSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr().setFitToPage(true);
      scalingSheet.getCTWorksheet().addNewPageSetup().setFitToWidth((short) 2);
      scalingSheet.getCTWorksheet().getPageSetup().setFitToHeight((short) 3);
      ExcelPrintLayoutController.applyScaling(
          scalingSheet, new ExcelPrintLayout.Scaling.Automatic());
      assertFalse(ExcelPrintLayoutController.pageSetupOrNull(scalingSheet).isSetFitToWidth());
      assertFalse(ExcelPrintLayoutController.pageSetupOrNull(scalingSheet).isSetFitToHeight());

      var automaticSheet = workbook.createSheet("Automatic");
      assertEquals(
          new ExcelPrintLayout.Scaling.Automatic(),
          ExcelPrintLayoutController.scaling(automaticSheet));
      automaticSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr();
      assertEquals(
          new ExcelPrintLayout.Scaling.Automatic(),
          ExcelPrintLayoutController.scaling(automaticSheet));

      var fitWithoutPageSetup = workbook.createSheet("FitWithoutPageSetup");
      fitWithoutPageSetup.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr().setFitToPage(true);
      assertEquals(
          new ExcelPrintLayout.Scaling.Fit(1, 1),
          ExcelPrintLayoutController.scaling(fitWithoutPageSetup));

      var fitSheet = workbook.createSheet("Fit");
      ExcelPrintLayoutController.applyScaling(fitSheet, new ExcelPrintLayout.Scaling.Fit(2, 3));
      assertTrue(fitSheet.getCTWorksheet().isSetPageSetup());
      assertEquals(
          new ExcelPrintLayout.Scaling.Fit(2, 3), ExcelPrintLayoutController.scaling(fitSheet));
    }
  }

  @Test
  void setPrintLayoutReplacesExistingRowAndColumnBreaks() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      ExcelPrintLayoutController controller = new ExcelPrintLayoutController();
      sheet.setRowBreak(1);
      sheet.setRowBreak(3);
      sheet.setColumnBreak(2);

      controller.setPrintLayout(
          sheet,
          new ExcelPrintLayout(
              new ExcelPrintLayout.Area.None(),
              ExcelPrintOrientation.PORTRAIT,
              new ExcelPrintLayout.Scaling.Automatic(),
              new ExcelPrintLayout.TitleRows.None(),
              new ExcelPrintLayout.TitleColumns.None(),
              ExcelHeaderFooterText.blank(),
              ExcelHeaderFooterText.blank(),
              new ExcelPrintSetup(
                  ExcelPrintSetup.defaults().margins(),
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
      ExcelPrintLayoutController.normalizeHeaderFooterNode(other);
      blankHeaderFooterSheet.getHeader().setLeft("");
      blankHeaderFooterSheet.getFooter().setRight("");
      assertTrue(blankHeaderFooterSheet.getCTWorksheet().isSetHeaderFooter());
      ExcelPrintLayoutController.normalizeHeaderFooterNode(blankHeaderFooterSheet);
      assertFalse(blankHeaderFooterSheet.getCTWorksheet().isSetHeaderFooter());
      var nonBlankFooterSheet = workbook.createSheet("NonBlankFooter");
      nonBlankFooterSheet.getFooter().setRight("keep");
      ExcelPrintLayoutController.normalizeHeaderFooterNode(nonBlankFooterSheet);
      assertTrue(nonBlankFooterSheet.getCTWorksheet().isSetHeaderFooter());
    }
  }

  @Test
  void exposesInternalPageSetupHelpersForNormalizationEdgeStates() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertPageSetupOrientationNormalization(workbook);
      assertEmptyAndNonEmptyPageSetupStates(workbook);
      assertPageSetupPropertiesNormalization(workbook);
    }
  }

  private static void assertPageSetupOrientationNormalization(XSSFWorkbook workbook) {
    var pageSetupSheet = workbook.createSheet("PageSetup");
    var pageSetup = pageSetupSheet.getCTWorksheet().addNewPageSetup();
    pageSetup.setOrientation(STOrientation.PORTRAIT);
    assertTrue(
        ExcelPrintLayoutController.shouldUnsetPageSetupOrientation(pageSetupSheet, pageSetup));
    ExcelPrintLayoutController.normalizePageSetupNode(pageSetupSheet);
    assertFalse(pageSetupSheet.getCTWorksheet().isSetPageSetup());

    var noPageSetupSheet = workbook.createSheet("NoPageSetup");
    ExcelPrintLayoutController.normalizePageSetupNode(noPageSetupSheet);
    assertNull(ExcelPrintLayoutController.pageSetupOrNull(noPageSetupSheet));

    var fitWidthSheet = workbook.createSheet("FitWidth");
    var fitWidthPageSetup = fitWidthSheet.getCTWorksheet().addNewPageSetup();
    fitWidthPageSetup.setOrientation(STOrientation.PORTRAIT);
    fitWidthPageSetup.setFitToWidth((short) 1);
    assertFalse(
        ExcelPrintLayoutController.shouldUnsetPageSetupOrientation(
            fitWidthSheet, fitWidthPageSetup));

    var fitHeightSheet = workbook.createSheet("FitHeight");
    var fitHeightPageSetup = fitHeightSheet.getCTWorksheet().addNewPageSetup();
    fitHeightPageSetup.setOrientation(STOrientation.PORTRAIT);
    fitHeightPageSetup.setFitToHeight((short) 1);
    assertFalse(
        ExcelPrintLayoutController.shouldUnsetPageSetupOrientation(
            fitHeightSheet, fitHeightPageSetup));

    var noOrientationSheet = workbook.createSheet("NoOrientation");
    var noOrientationPageSetup = noOrientationSheet.getCTWorksheet().addNewPageSetup();
    assertFalse(
        ExcelPrintLayoutController.shouldUnsetPageSetupOrientation(
            noOrientationSheet, noOrientationPageSetup));
  }

  private static void assertEmptyAndNonEmptyPageSetupStates(XSSFWorkbook workbook) {
    var fitWidthOnlySheet = workbook.createSheet("FitWidthOnly");
    var fitWidthOnlyPageSetup = fitWidthOnlySheet.getCTWorksheet().addNewPageSetup();
    fitWidthOnlyPageSetup.setFitToWidth((short) 1);
    assertFalse(ExcelPrintLayoutController.isEmptyPageSetup(fitWidthOnlyPageSetup));

    var fitHeightOnlySheet = workbook.createSheet("FitHeightOnly");
    var fitHeightOnlyPageSetup = fitHeightOnlySheet.getCTWorksheet().addNewPageSetup();
    fitHeightOnlyPageSetup.setFitToHeight((short) 1);
    assertFalse(ExcelPrintLayoutController.isEmptyPageSetup(fitHeightOnlyPageSetup));

    var printerDefaultsSheet = workbook.createSheet("PrinterDefaults");
    var printerDefaults = printerDefaultsSheet.getCTWorksheet().addNewPageSetup();
    printerDefaults.setUsePrinterDefaults(true);
    assertFalse(ExcelPrintLayoutController.isEmptyPageSetup(printerDefaults));

    var advancedSetupSheet = workbook.createSheet("AdvancedSetup");
    var advancedSetup = advancedSetupSheet.getCTWorksheet().addNewPageSetup();
    advancedSetup.setPaperSize((short) 8);
    advancedSetup.setDraft(true);
    advancedSetup.setBlackAndWhite(true);
    advancedSetup.setCopies((short) 2);
    advancedSetup.setUseFirstPageNumber(true);
    advancedSetup.setFirstPageNumber((short) 4);
    assertFalse(ExcelPrintLayoutController.isEmptyPageSetup(advancedSetup));
  }

  private static void assertPageSetupPropertiesNormalization(XSSFWorkbook workbook) {
    var missingSheetPr = workbook.createSheet("MissingSheetPr");
    ExcelPrintLayoutController.normalizePageSetupProperties(missingSheetPr);
    assertNull(ExcelPrintLayoutController.sheetPrOrNull(missingSheetPr));

    var staleSheetPr = workbook.createSheet("StaleSheetPr");
    staleSheetPr.getCTWorksheet().addNewSheetPr();
    ExcelPrintLayoutController.normalizePageSetupProperties(staleSheetPr);
    assertFalse(staleSheetPr.getCTWorksheet().isSetSheetPr());
    assertNull(ExcelPrintLayoutController.pageSetUpPrOrNull(staleSheetPr));

    var unsetFitSheet = workbook.createSheet("UnsetFit");
    unsetFitSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr();
    ExcelPrintLayoutController.normalizePageSetupProperties(unsetFitSheet);
    assertFalse(unsetFitSheet.getCTWorksheet().isSetSheetPr());

    var falseFitSheet = workbook.createSheet("FalseFit");
    falseFitSheet.getCTWorksheet().addNewSheetPr().addNewPageSetUpPr().setFitToPage(false);
    ExcelPrintLayoutController.normalizePageSetupProperties(falseFitSheet);
    assertFalse(falseFitSheet.getCTWorksheet().isSetSheetPr());
  }

  private static ExcelPrintLayout factualDefaultPrintLayout() {
    return new ExcelPrintLayout(
        new ExcelPrintLayout.Area.None(),
        ExcelPrintOrientation.PORTRAIT,
        new ExcelPrintLayout.Scaling.Automatic(),
        new ExcelPrintLayout.TitleRows.None(),
        new ExcelPrintLayout.TitleColumns.None(),
        ExcelHeaderFooterText.blank(),
        ExcelHeaderFooterText.blank(),
        new ExcelPrintSetup(
            new ExcelPrintMargins(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
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
