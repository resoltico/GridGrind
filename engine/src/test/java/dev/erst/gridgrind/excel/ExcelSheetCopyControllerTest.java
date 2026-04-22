package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

/** Tests for sheet-copy planning and validation helpers. */
class ExcelSheetCopyControllerTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void copySheetPreservesProtectionOnEmptySourceSheets() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Source");
      workbook.setSheetProtection("Source", protectionSettings());

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      assertEquals(List.of("Source", "Replica"), workbook.sheetNames());
      assertEquals(
          new WorkbookReadResult.SheetProtection.Protected(protectionSettings()),
          workbook.sheetSummary("Replica").protection());
      assertEquals(0, workbook.sheetSummary("Replica").physicalRowCount());
    }
  }

  @Test
  void copySheetSupportsReopenedPoiCommentWorkbooks() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-comments-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Source");
      sheet.createRow(0).createCell(0).setCellValue("Lead");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFClientAnchor anchor = drawing.createAnchor(64, 24, 448, 96, 0, 0, 3, 3);
      XSSFComment comment = drawing.createCellComment(anchor);
      comment.setString(new XSSFRichTextString("Review"));
      comment.setAuthor("GridGrind");
      sheet.getRow(0).getCell(0).setCellComment(comment);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertDoesNotThrow(
          () -> workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd()));
      assertEquals(
          "Review",
          workbook
              .sheet("Replica")
              .comments(new ExcelCellSelection.AllUsedCells())
              .getFirst()
              .comment()
              .text());
    }
  }

  @Test
  void copySheetPreservesSheetProtectionPasswordHashes() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Source");
      workbook.setSheetProtection("Source", protectionSettings(), "gridgrind-copy");

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      assertTrue(
          workbook.xssfWorkbook().getSheet("Replica").validateSheetPassword("gridgrind-copy"));
      assertFalse(
          workbook.xssfWorkbook().getSheet("Replica").validateSheetPassword("gridgrind-wrong"));
      assertEquals(
          new WorkbookReadResult.SheetProtection.Protected(protectionSettings()),
          workbook.sheetSummary("Replica").protection());
    }
  }

  @Test
  void copySheetPreservesAdvancedPrintSetupWhenOnlyPageSetupAttrsAreNonDefault()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Source");
      ExcelPrintLayout advancedPrintLayout =
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
                  List.of(6),
                  List.of(3)));
      workbook.sheet("Source").setPrintLayout(advancedPrintLayout);

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      assertEquals(advancedPrintLayout, workbook.sheet("Replica").printLayout());
    }
  }

  @Test
  void copySheetPreservesDrawingObjectsAndRetargetsCopiedCharts() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-drawings-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet source = workbook.getOrCreateSheet("Source");
      ExcelChartTestSupport.seedChartData(source);
      source.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              ExcelChartTestSupport.anchor(0, 5, 3, 9),
              "Queue preview"));
      source.setChart(
          ExcelChartTestSupport.barChart(
              "OpsChart",
              ExcelChartTestSupport.anchor(6, 1, 12, 14),
              new ExcelChartDefinition.Title.Text("Roadmap"),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              false,
              ExcelChartBarDirection.COLUMN,
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Text("Plan"),
                      ExcelChartTestSupport.ref("A2:A4"),
                      ExcelChartTestSupport.ref("B2:B4")))));

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      ExcelSheet replica = workbook.sheet("Replica");
      assertEquals(
          List.of("OpsPicture", "OpsChart"),
          replica.drawingObjects().stream().map(ExcelDrawingObjectSnapshot::name).toList());
      ExcelDrawingObjectSnapshot.Picture copiedPicture =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Picture.class,
              replica.drawingObjects().stream()
                  .filter(snapshot -> "OpsPicture".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(ExcelChartTestSupport.anchor(0, 5, 3, 9), copiedPicture.anchor());

      ExcelChartSnapshot copiedChart = replica.charts().getFirst();
      assertEquals("OpsChart", copiedChart.name());
      assertEquals(ExcelChartTestSupport.anchor(6, 1, 12, 14), copiedChart.anchor());
      ExcelChartSnapshot.Series copiedSeries =
          ExcelChartTestSupport.singlePlot(copiedChart, ExcelChartSnapshot.Bar.class)
              .series()
              .getFirst();
      assertEquals(
          "Replica!$A$2:$A$4",
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.StringReference.class, copiedSeries.categories())
              .formula());
      assertEquals(
          "Replica!$B$2:$B$4",
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.NumericReference.class, copiedSeries.values())
              .formula());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet replica = reopened.sheet("Replica");
      assertEquals(
          List.of("OpsPicture", "OpsChart"),
          replica.drawingObjects().stream().map(ExcelDrawingObjectSnapshot::name).toList());
      ExcelChartSnapshot copiedChart = replica.charts().getFirst();
      ExcelChartSnapshot.Series copiedSeries =
          ExcelChartTestSupport.singlePlot(copiedChart, ExcelChartSnapshot.Bar.class)
              .series()
              .getFirst();
      assertEquals(
          "Replica!$A$2:$A$4",
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.StringReference.class, copiedSeries.categories())
              .formula());
    }
  }

  @Test
  void copySheetRetargetsAdvancedWorkbookCoreStructures() throws IOException {
    ExcelAutofilterController autofilterController = new ExcelAutofilterController();
    ExcelTableController tableController = new ExcelTableController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet source = seedAdvancedCopySource(workbook);
      List<WorkbookReadResult.CellComment> sourceComments =
          source.comments(new ExcelCellSelection.Selected(List.of("E2")));

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      ExcelSheet replica = workbook.sheet("Replica");
      assertEquals("SUM(Replica!B2:B3)", replica.formula("C2"));
      assertEquals(
          sourceComments, replica.comments(new ExcelCellSelection.Selected(List.of("E2"))));
      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned(
                  "A1:F3",
                  List.of(
                      new ExcelAutofilterFilterColumnSnapshot(
                          0L,
                          false,
                          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Ada"), true)),
                      new ExcelAutofilterFilterColumnSnapshot(
                          1L,
                          true,
                          new ExcelAutofilterFilterCriterionSnapshot.Custom(
                              true,
                              List.of(
                                  new ExcelAutofilterFilterCriterionSnapshot.CustomCondition(
                                      "greaterThan", "1"),
                                  new ExcelAutofilterFilterCriterionSnapshot.CustomCondition(
                                      "equal", "Queue")))),
                      new ExcelAutofilterFilterColumnSnapshot(
                          2L,
                          true,
                          new ExcelAutofilterFilterCriterionSnapshot.Dynamic("today", 1.0d, 2.0d)),
                      new ExcelAutofilterFilterColumnSnapshot(
                          3L,
                          true,
                          new ExcelAutofilterFilterCriterionSnapshot.Top10(
                              false, true, 10.0d, null)),
                      new ExcelAutofilterFilterColumnSnapshot(
                          4L,
                          true,
                          new ExcelAutofilterFilterCriterionSnapshot.Color(
                              false, new ExcelColorSnapshot("#AABBCC"))),
                      new ExcelAutofilterFilterColumnSnapshot(
                          5L,
                          true,
                          new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", 2))),
                  new ExcelAutofilterSortStateSnapshot(
                      "A1:F3",
                      true,
                      true,
                      "stroke",
                      List.of(
                          new ExcelAutofilterSortConditionSnapshot(
                              "A2:A3", true, "cellColor", new ExcelColorSnapshot("#102030"), null),
                          new ExcelAutofilterSortConditionSnapshot(
                              "B2:B3", false, "icon", null, 4))))),
          autofilterController.sheetOwnedAutofilters(replica.xssfSheet()));
      List<ExcelDataValidationSnapshot> replicaValidations =
          replica.dataValidations(new ExcelRangeSelection.All());
      assertEquals(2, replicaValidations.size());
      ExcelDataValidationSnapshot.Supported formulaList =
          assertInstanceOf(ExcelDataValidationSnapshot.Supported.class, replicaValidations.get(0));
      assertEquals(
          new ExcelDataValidationRule.FormulaList("Replica!$J$2:$J$3"),
          formulaList.validation().rule());
      ExcelDataValidationSnapshot.Supported wholeNumber =
          assertInstanceOf(ExcelDataValidationSnapshot.Supported.class, replicaValidations.get(1));
      assertInstanceOf(ExcelDataValidationRule.WholeNumber.class, wholeNumber.validation().rule());
      List<CTDataValidation> rawReplicaValidations =
          List.of(
              replica.xssfSheet().getCTWorksheet().getDataValidations().getDataValidationArray());
      assertEquals("Replica!$J$2:$J$3", rawReplicaValidations.get(0).getFormula1());
      assertEquals("1", rawReplicaValidations.get(1).getFormula1());
      assertTrue(rawReplicaValidations.get(1).getFormula2().contains("Replica"));
      assertTrue(rawReplicaValidations.get(1).getFormula2().contains("$B$2:$B$3"));
      List<ExcelNamedRangeSnapshot> replicaNames =
          workbook.namedRanges().stream()
              .filter(
                  namedRange ->
                      namedRange.scope() instanceof ExcelNamedRangeScope.SheetScope scope
                          && "Replica".equals(scope.sheetName()))
              .toList();
      assertEquals(2, replicaNames.size());
      assertTrue(
          replicaNames.stream()
              .anyMatch(
                  namedRange ->
                      namedRange instanceof ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot
                          && "LocalRange".equals(rangeSnapshot.name())
                          && rangeSnapshot.refersToFormula().contains("Replica")
                          && rangeSnapshot.refersToFormula().contains("$A$2")
                          && rangeSnapshot.refersToFormula().contains("$A$3")
                          && "A2:A3".equals(rangeSnapshot.target().range())));
      assertTrue(
          replicaNames.stream()
              .anyMatch(
                  namedRange ->
                      namedRange instanceof ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot
                          && "LocalFormula".equals(formulaSnapshot.name())
                          && formulaSnapshot.refersToFormula().contains("Replica")
                          && formulaSnapshot.refersToFormula().contains("$B$2:$B$3")));
      ExcelConditionalFormattingBlockSnapshot copiedFormatting =
          replica.conditionalFormatting(new ExcelRangeSelection.All()).getFirst();
      assertEquals(
          "SUM(Replica!$B$2:$B$3)>0",
          assertInstanceOf(
                  ExcelConditionalFormattingRuleSnapshot.FormulaRule.class,
                  copiedFormatting.rules().get(0))
              .formula());
      assertEquals(
          "SUM(Replica!$B$2:$B$3)",
          assertInstanceOf(
                  ExcelConditionalFormattingRuleSnapshot.CellValueRule.class,
                  copiedFormatting.rules().get(1))
              .formula2());
      ExcelTableSnapshot copiedTable =
          tableController.tables(workbook, new ExcelTableSelection.All()).stream()
              .filter(table -> "Replica".equals(table.sheetName()))
              .findFirst()
              .orElseThrow();
      assertEquals("L1:M3", copiedTable.range());
      assertNotEquals("ReplicaTable", copiedTable.name());
      assertEquals("replica table", copiedTable.comment());
      assertTrue(copiedTable.published());
      assertTrue(copiedTable.insertRow());
      assertFalse(copiedTable.insertRowShift());
      assertTrue(
          workbook.xssfWorkbook().getSheet("Replica").validateSheetPassword("gridgrind-copy"));
    }
  }

  @Test
  void copySheetPreservesAdvancedCommentsAfterSaveAndReopen() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-comment-reopen-");

    List<WorkbookReadResult.CellComment> sourceComments;
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet source = seedAdvancedCopySource(workbook);
      sourceComments = source.comments(new ExcelCellSelection.Selected(List.of("E2")));

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());
      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertEquals(
          sourceComments,
          reopened.sheet("Replica").comments(new ExcelCellSelection.Selected(List.of("E2"))));
    }
  }

  @Test
  void copySheetRetargetsSupportedConditionalFormattingFamiliesAndNullSortState()
      throws IOException {
    ExcelAutofilterController autofilterController = new ExcelAutofilterController();
    ExcelTableController tableController = new ExcelTableController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet source = workbook.getOrCreateSheet("Source");
      source.setCell("A1", ExcelCellValue.text("Owner"));
      source.setCell("B1", ExcelCellValue.text("Amount"));
      source.setCell("C1", ExcelCellValue.text("Flag"));
      source.setCell("A2", ExcelCellValue.text("Ada"));
      source.setCell("A3", ExcelCellValue.text("Lin"));
      source.setCell("B2", ExcelCellValue.number(2));
      source.setCell("B3", ExcelCellValue.number(4));
      source.setCell("E1", ExcelCellValue.text("Desk"));
      source.setCell("F1", ExcelCellValue.text("Region"));
      source.setCell("E2", ExcelCellValue.text("A1"));
      source.setCell("E3", ExcelCellValue.text("B1"));
      source.setCell("E4", ExcelCellValue.text("Totals"));
      source.setCell("F2", ExcelCellValue.text("North"));
      source.setCell("F3", ExcelCellValue.text("South"));
      source.setCell("F4", ExcelCellValue.text("2"));

      source.setAutofilter(
          "A1:B3",
          List.of(
              new ExcelAutofilterFilterColumn(
                  0L, true, new ExcelAutofilterFilterCriterion.Values(List.of("Ada"), false))),
          null);
      var rawValidations = source.xssfSheet().getCTWorksheet().addNewDataValidations();
      CTDataValidation rawValidation = rawValidations.addNewDataValidation();
      rawValidation.setSqref(List.of("C2:C3"));
      rawValidation.setFormula1("SUM(Source!$B$2:$B$3)");
      rawValidations.setCount(rawValidations.sizeOfDataValidationArray());
      source.setConditionalFormatting(
          new ExcelConditionalFormattingBlockDefinition(
              List.of("C2:C3"),
              List.of(
                  new ExcelConditionalFormattingRule.FormulaRule(
                      "SUM(Source!$B$2:$B$3)>0", false, null),
                  new ExcelConditionalFormattingRule.CellValueRule(
                      ExcelComparisonOperator.GREATER_THAN,
                      "SUM(Source!$B$2:$B$3)",
                      null,
                      false,
                      null),
                  new ExcelConditionalFormattingRule.ColorScaleRule(
                      List.of(
                          new ExcelConditionalFormattingThreshold(
                              ExcelConditionalFormattingThresholdType.MIN, null, null),
                          new ExcelConditionalFormattingThreshold(
                              ExcelConditionalFormattingThresholdType.MAX, null, null)),
                      List.of(new ExcelColor("#112233"), new ExcelColor("#445566")),
                      false),
                  new ExcelConditionalFormattingRule.DataBarRule(
                      new ExcelColor("#223344"),
                      false,
                      0,
                      100,
                      new ExcelConditionalFormattingThreshold(
                          ExcelConditionalFormattingThresholdType.MIN, null, null),
                      new ExcelConditionalFormattingThreshold(
                          ExcelConditionalFormattingThresholdType.MAX, null, null),
                      false),
                  new ExcelConditionalFormattingRule.IconSetRule(
                      ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                      false,
                      true,
                      List.of(
                          new ExcelConditionalFormattingThreshold(
                              ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                          new ExcelConditionalFormattingThreshold(
                              ExcelConditionalFormattingThresholdType.PERCENT, null, 50.0d),
                          new ExcelConditionalFormattingThreshold(
                              ExcelConditionalFormattingThresholdType.MAX, null, null)),
                      false),
                  new ExcelConditionalFormattingRule.Top10Rule(5, false, false, false, null))));
      workbook.setTable(
          new ExcelTableDefinition(
              "Queue",
              "Source",
              "E1:F4",
              true,
              false,
              new ExcelTableStyle.None(),
              "",
              false,
              false,
              false,
              "",
              "",
              "",
              List.of()));

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned(
                  "A1:B3",
                  List.of(
                      new ExcelAutofilterFilterColumnSnapshot(
                          0L,
                          true,
                          new ExcelAutofilterFilterCriterionSnapshot.Values(
                              List.of("Ada"), false))),
                  null)),
          autofilterController.sheetOwnedAutofilters(workbook.sheet("Replica").xssfSheet()));
      assertEquals(
          "SUM(Replica!$B$2:$B$3)",
          workbook
              .sheet("Replica")
              .xssfSheet()
              .getCTWorksheet()
              .getDataValidations()
              .getDataValidationArray(0)
              .getFormula1());

      ExcelConditionalFormattingBlockSnapshot copiedFormatting =
          workbook.sheet("Replica").conditionalFormatting(new ExcelRangeSelection.All()).getFirst();
      assertEquals(
          "SUM(Replica!$B$2:$B$3)>0",
          assertInstanceOf(
                  ExcelConditionalFormattingRuleSnapshot.FormulaRule.class,
                  copiedFormatting.rules().get(0))
              .formula());
      ExcelConditionalFormattingRuleSnapshot.CellValueRule copiedCellValueRule =
          assertInstanceOf(
              ExcelConditionalFormattingRuleSnapshot.CellValueRule.class,
              copiedFormatting.rules().get(1));
      assertEquals("SUM(Replica!$B$2:$B$3)", copiedCellValueRule.formula1());
      assertNull(copiedCellValueRule.formula2());
      assertInstanceOf(
          ExcelConditionalFormattingRuleSnapshot.ColorScaleRule.class,
          copiedFormatting.rules().get(2));
      assertInstanceOf(
          ExcelConditionalFormattingRuleSnapshot.DataBarRule.class,
          copiedFormatting.rules().get(3));
      ExcelConditionalFormattingRuleSnapshot.IconSetRule copiedIconSetRule =
          assertInstanceOf(
              ExcelConditionalFormattingRuleSnapshot.IconSetRule.class,
              copiedFormatting.rules().get(4));
      assertTrue(copiedIconSetRule.reversed());
      assertInstanceOf(
          ExcelConditionalFormattingRuleSnapshot.Top10Rule.class, copiedFormatting.rules().get(5));

      ExcelTableSnapshot copiedTable =
          tableController.tables(workbook, new ExcelTableSelection.All()).stream()
              .filter(table -> "Replica".equals(table.sheetName()))
              .findFirst()
              .orElseThrow();
      assertEquals("E1:F4", copiedTable.range());
      assertEquals(1, copiedTable.totalsRowCount());
      assertFalse(copiedTable.hasAutofilter());
      assertInstanceOf(ExcelTableStyleSnapshot.None.class, copiedTable.style());
    }
  }

  @Test
  void requireNoTablesIsNowANoOpBecauseSheetCopySupportsTables() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Source");
      sheet.createRow(0).createCell(0).setCellValue("Header");
      sheet.getRow(0).createCell(1).setCellValue("Value");
      sheet.createRow(1).createCell(0).setCellValue("A");
      sheet.getRow(1).createCell(1).setCellValue("B");
      sheet.createRow(2).createCell(0).setCellValue("C");
      sheet.getRow(2).createCell(1).setCellValue("D");
      sheet.createTable(
          new org.apache.poi.ss.util.AreaReference(
              "A1:B3", org.apache.poi.ss.SpreadsheetVersion.EXCEL2007));

      assertDoesNotThrow(() -> ExcelSheetCopyController.requireNoTables(sheet, "Source"));
    }
  }

  @Test
  void localNamedRangeHelpersRespectScopeBoundaries() {
    ExcelNamedRangeSnapshot.FormulaSnapshot localFormula =
        new ExcelNamedRangeSnapshot.FormulaSnapshot(
            "LocalFormula", new ExcelNamedRangeScope.SheetScope("Source"), "SUM(Source!$A$1:$A$2)");
    List<ExcelNamedRangeSnapshot> namedRanges =
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "WorkbookBudget",
                new ExcelNamedRangeScope.WorkbookScope(),
                "Source!$A$1",
                new ExcelNamedRangeTarget("Source", "A1")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "LocalBudget",
                new ExcelNamedRangeScope.SheetScope("Source"),
                "Source!$A$1:$A$2",
                new ExcelNamedRangeTarget("Source", "A1:A2")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "OtherLocal",
                new ExcelNamedRangeScope.SheetScope("Other"),
                "Other!$B$1",
                new ExcelNamedRangeTarget("Other", "B1")),
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "OtherFormula", new ExcelNamedRangeScope.SheetScope("Other"), "SUM(Other!$A$1)"),
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "WorkbookFormula", new ExcelNamedRangeScope.WorkbookScope(), "SUM(Source!$A$1)"),
            localFormula);

    assertEquals(
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "LocalBudget",
                new ExcelNamedRangeScope.SheetScope("Source"),
                "Source!$A$1:$A$2",
                new ExcelNamedRangeTarget("Source", "A1:A2"))),
        ExcelSheetCopyController.copyableLocalRangeNames(namedRanges, "Source"));
    assertDoesNotThrow(
        () -> ExcelSheetCopyController.requireNoUncopyableLocalNamedRanges(namedRanges, "Source"));
    assertEquals(
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "LocalBudget",
                new ExcelNamedRangeScope.SheetScope("Source"),
                "Source!$A$1:$A$2",
                new ExcelNamedRangeTarget("Source", "A1:A2")),
            localFormula),
        ExcelSheetCopyController.copyableLocalNames(namedRanges, "Source"));
  }

  @Test
  void copySheetPreservesAndRetargetsRawDataValidations() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Source");
      workbook
          .sheet("Source")
          .setDataValidation(
              "A1:A3",
              new ExcelDataValidationDefinition(
                  new ExcelDataValidationRule.FormulaList("Source!$D$1:$D$3"),
                  false,
                  false,
                  null,
                  null));
      addRawValidation(
          workbook.xssfWorkbook().getSheet("Source"), "B1:B3", STDataValidationType.LIST, "\"\"");
      addRawValidation(
          workbook.xssfWorkbook().getSheet("Source"), "C1:C3", STDataValidationType.LIST, null);

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      List<ExcelDataValidationSnapshot> replicaValidations =
          workbook.sheet("Replica").dataValidations(new ExcelRangeSelection.All());
      assertEquals(3, replicaValidations.size());
      ExcelDataValidationSnapshot.Supported formulaList =
          assertInstanceOf(ExcelDataValidationSnapshot.Supported.class, replicaValidations.get(0));
      assertEquals(
          new ExcelDataValidationRule.FormulaList("Replica!$D$1:$D$3"),
          formulaList.validation().rule());
      ExcelDataValidationSnapshot.Supported emptyExplicitList =
          assertInstanceOf(ExcelDataValidationSnapshot.Supported.class, replicaValidations.get(1));
      assertEquals(
          new ExcelDataValidationRule.ExplicitList(List.of()),
          emptyExplicitList.validation().rule());
      assertEquals(
          new ExcelDataValidationSnapshot.Unsupported(
              List.of("C1:C3"),
              "MISSING_FORMULA",
              "List validation is missing both explicit values and formula1."),
          replicaValidations.get(2));
      List<CTDataValidation> rawReplicaValidations =
          List.of(
              workbook
                  .xssfWorkbook()
                  .getSheet("Replica")
                  .getCTWorksheet()
                  .getDataValidations()
                  .getDataValidationArray());
      assertEquals("Replica!$D$1:$D$3", rawReplicaValidations.get(0).getFormula1());
      assertEquals("\"\"", rawReplicaValidations.get(1).getFormula1());
      assertFalse(rawReplicaValidations.get(2).isSetFormula1());
    }
  }

  @Test
  void copySheetHandlesEdgeCasesInValidationFormulaRetargeting() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Source");
      var sheet = workbook.xssfWorkbook().getSheet("Source");
      // List validation: formula1.length() < 2 — covers isQuotedListLiteral length < 2 branch.
      addRawValidation(sheet, "A1", STDataValidationType.LIST, "1");
      // List validation: starts with '"' but does not end with '"' — covers startsWith=true,
      // endsWith=false branch in isQuotedListLiteral (valid formula, not a quoted literal).
      addRawValidation(sheet, "B1", STDataValidationType.LIST, "\"prefix\"&Source!A1");
      // Non-list validation: no formula1 set — covers isSetFormula1()=false branch.
      addRawValidation(sheet, "C1", STDataValidationType.WHOLE, null);
      // Non-list validation: formula1 is blank — covers isSetFormula1()=true, isBlank()=true
      // branch.
      addRawValidation(sheet, "D1", STDataValidationType.WHOLE, " ");
      // Non-list validation: formula1 non-blank, formula2 not set — covers isSetFormula2()=false.
      addRawValidation(sheet, "E1", STDataValidationType.WHOLE, "1");
      // Non-list validation: formula1 non-blank, formula2 blank — covers isSetFormula2()=true,
      // isBlank()=true branch.
      addRawValidationWithFormulas(sheet, "F1", STDataValidationType.WHOLE, "1", " ");
      // Raw validation without a declared type still needs formula retargeting.
      var validations = sheet.getCTWorksheet().getDataValidations();
      CTDataValidation missingType = validations.addNewDataValidation();
      missingType.setSqref(List.of("G1"));
      missingType.setFormula1("Source!$A$1");
      validations.setCount(validations.sizeOfDataValidationArray());

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      List<CTDataValidation> rawReplica =
          List.of(
              workbook
                  .xssfWorkbook()
                  .getSheet("Replica")
                  .getCTWorksheet()
                  .getDataValidations()
                  .getDataValidationArray());
      assertEquals(7, rawReplica.size());
      assertEquals("1", rawReplica.get(0).getFormula1());
      assertTrue(rawReplica.get(1).getFormula1().contains("Replica"));
      assertFalse(rawReplica.get(2).isSetFormula1());
      assertEquals(" ", rawReplica.get(3).getFormula1());
      assertEquals("1", rawReplica.get(4).getFormula1());
      assertFalse(rawReplica.get(4).isSetFormula2());
      assertEquals("1", rawReplica.get(5).getFormula1());
      assertEquals(" ", rawReplica.get(5).getFormula2());
      assertEquals("Replica!$A$1", rawReplica.get(6).getFormula1());
    }
  }

  @Test
  void supportedConditionalFormattingCopiesEverySupportedRuleFamily() {
    ExcelConditionalFormattingBlockSnapshot supportedBlock =
        new ExcelConditionalFormattingBlockSnapshot(
            List.of("A1:A2"),
            List.of(
                new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                    1, true, "A1>0", supportedStyle()),
                new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
                    2, false, ExcelComparisonOperator.GREATER_THAN, "0", null, supportedStyle())));

    assertEquals(
        1,
        ExcelSheetCopyController.supportedConditionalFormatting(List.of(supportedBlock), "Source")
            .size());
    assertInstanceOf(
        ExcelConditionalFormattingRule.ColorScaleRule.class,
        ExcelSheetCopyController.copyableRule(
            new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
                1,
                false,
                List.of(
                    new ExcelConditionalFormattingThresholdSnapshot(
                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                    new ExcelConditionalFormattingThresholdSnapshot(
                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                List.of("#102030", "#405060")),
            "Source"));
    assertInstanceOf(
        ExcelConditionalFormattingRule.DataBarRule.class,
        ExcelSheetCopyController.copyableRule(
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                1,
                false,
                "#102030",
                false,
                0,
                100,
                new ExcelConditionalFormattingThresholdSnapshot(
                    ExcelConditionalFormattingThresholdType.MIN, null, null),
                new ExcelConditionalFormattingThresholdSnapshot(
                    ExcelConditionalFormattingThresholdType.MAX, null, null)),
            "Source"));
    assertInstanceOf(
        ExcelConditionalFormattingRule.IconSetRule.class,
        ExcelSheetCopyController.copyableRule(
            new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
                1,
                false,
                ExcelConditionalFormattingIconSet.GYR_3_ARROW,
                false,
                false,
                List.of(
                    new ExcelConditionalFormattingThresholdSnapshot(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                    new ExcelConditionalFormattingThresholdSnapshot(
                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                    new ExcelConditionalFormattingThresholdSnapshot(
                        ExcelConditionalFormattingThresholdType.MAX, null, null))),
            "Source"));
    assertInstanceOf(
        ExcelConditionalFormattingRule.Top10Rule.class,
        ExcelSheetCopyController.copyableRule(
            new ExcelConditionalFormattingRuleSnapshot.Top10Rule(
                1, false, 10, true, false, supportedStyle()),
            "Source"));
    assertInstanceOf(
        ExcelConditionalFormattingRule.FormulaRule.class,
        ExcelSheetCopyController.copyableRule(
            new ExcelConditionalFormattingRuleSnapshot.FormulaRule(1, false, "A1>0", null),
            "Source"));
    assertInstanceOf(
        ExcelConditionalFormattingRule.CellValueRule.class,
        ExcelSheetCopyController.copyableRule(
            new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
                1, false, ExcelComparisonOperator.GREATER_THAN, "0", null, null),
            "Source"));
    assertInstanceOf(
        ExcelConditionalFormattingRule.Top10Rule.class,
        ExcelSheetCopyController.copyableRule(
            new ExcelConditionalFormattingRuleSnapshot.Top10Rule(1, false, 10, false, false, null),
            "Source"));
    assertEquals(
        "cannot copy sheet 'Source': unsupported conditional-formatting rule 'TOP10' is not copyable",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    ExcelSheetCopyController.copyableRule(
                        new ExcelConditionalFormattingRuleSnapshot.UnsupportedRule(
                            1, false, "TOP10", "detail"),
                        "Source"))
            .getMessage());
  }

  @Test
  void copyableStyleRejectsUnsupportedDifferentialStyleFeatures() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelSheetCopyController.copyableStyle(
                    new ExcelDifferentialStyleSnapshot(
                        "0.00",
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        null,
                        List.of(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT)),
                    "Source"));
    assertEquals(
        "cannot copy sheet 'Source': conditional-formatting rules with unsupported differential-style features are not copyable",
        failure.getMessage());
  }

  @Test
  void sheetOwnedAutofilterRangeHandlesEmptySheetOwnedAndInvalidTableOwnedStates() {
    assertEquals(Optional.empty(), ExcelSheetCopyController.sheetOwnedAutofilterRange(List.of()));
    assertEquals(
        Optional.of("A1:B3"),
        ExcelSheetCopyController.sheetOwnedAutofilterRange(
            List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3"))));
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelSheetCopyController.sheetOwnedAutofilterRange(
                    List.of(new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Ops"))));
    assertEquals(
        "sheetOwnedAutofilters must not return table-owned autofilter snapshots",
        failure.getMessage());
  }

  private static void addRawValidation(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      String range,
      STDataValidationType.Enum type,
      String formula1) {
    addRawValidationWithFormulas(sheet, range, type, formula1, null);
  }

  private static void addRawValidationWithFormulas(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      String range,
      STDataValidationType.Enum type,
      String formula1,
      String formula2) {
    var validations =
        sheet.getCTWorksheet().isSetDataValidations()
            ? sheet.getCTWorksheet().getDataValidations()
            : sheet.getCTWorksheet().addNewDataValidations();
    CTDataValidation validation = validations.addNewDataValidation();
    validation.setSqref(List.of(range));
    validation.setType(type);
    if (formula1 != null) {
      validation.setFormula1(formula1);
    }
    if (formula2 != null) {
      validation.setFormula2(formula2);
    }
    validations.setCount(validations.sizeOfDataValidationArray());
  }

  private static ExcelSheet seedAdvancedCopySource(ExcelWorkbook workbook) throws IOException {
    ExcelSheet source = workbook.getOrCreateSheet("Source");
    source.setCell("A1", ExcelCellValue.text("Owner"));
    source.setCell("B1", ExcelCellValue.text("Amount"));
    source.setCell("C1", ExcelCellValue.text("Formula"));
    source.setCell("D1", ExcelCellValue.text("Stage"));
    source.setCell("E1", ExcelCellValue.text("Comment"));
    source.setCell("F1", ExcelCellValue.text("Flag"));
    source.setCell("G1", ExcelCellValue.text("List"));
    source.setCell("H1", ExcelCellValue.text("Whole"));
    source.setCell("I1", ExcelCellValue.text("Note"));
    source.setCell("J1", ExcelCellValue.text("Statuses"));
    source.setCell("A2", ExcelCellValue.text("Ada"));
    source.setCell("A3", ExcelCellValue.text("Lin"));
    source.setCell("B2", ExcelCellValue.number(2));
    source.setCell("B3", ExcelCellValue.number(4));
    source.setCell("C2", ExcelCellValue.formula("SUM(Source!B2:B3)"));
    source.setCell("D2", ExcelCellValue.text("Today"));
    source.setCell("D3", ExcelCellValue.text("Today"));
    source.setCell("F2", ExcelCellValue.text("High"));
    source.setCell("F3", ExcelCellValue.text("Low"));
    source.setCell("J2", ExcelCellValue.text("Ready"));
    source.setCell("J3", ExcelCellValue.text("Done"));
    source.setCell("H2", ExcelCellValue.number(1));
    source.setCell("H3", ExcelCellValue.number(2));
    source.setCell("E2", ExcelCellValue.text("Quarterly Review"));
    source.setComment(
        "E2",
        new ExcelComment(
            "Quarterly Review",
            "GridGrind",
            true,
            new ExcelRichText(
                List.of(
                    new ExcelRichTextRun("Quarterly", null),
                    new ExcelRichTextRun(
                        " Review",
                        new ExcelCellFont(
                            true, false, "Aptos", null, new ExcelColor("#112233"), null, null)))),
            new ExcelCommentAnchor(1, 1, 4, 5)));
    source.setAutofilter("A1:F3", advancedAutofilterCriteria(), advancedAutofilterSortState());
    source.setDataValidation(
        "G2:G3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.FormulaList("Source!$J$2:$J$3"), false, false, null, null));
    source.setDataValidation(
        "H2:H3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.WholeNumber(
                ExcelComparisonOperator.BETWEEN, "1", "SUM(Source!$B$2:$B$3)"),
            false,
            false,
            null,
            null));
    source.setConditionalFormatting(
        new ExcelConditionalFormattingBlockDefinition(
            List.of("I2:I3"),
            List.of(
                new ExcelConditionalFormattingRule.FormulaRule(
                    "SUM(Source!$B$2:$B$3)>0",
                    false,
                    ExcelSheetCopyController.copyableStyle(supportedStyle(), "Source")),
                new ExcelConditionalFormattingRule.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    "SUM(Source!$B$2:$B$3)",
                    false,
                    ExcelSheetCopyController.copyableStyle(supportedStyle(), "Source")))));
    workbook.setNamedRange(
        new ExcelNamedRangeDefinition(
            "LocalRange",
            new ExcelNamedRangeScope.SheetScope("Source"),
            new ExcelNamedRangeTarget("Source", "A2:A3")));
    workbook.setNamedRange(
        new ExcelNamedRangeDefinition(
            "LocalFormula",
            new ExcelNamedRangeScope.SheetScope("Source"),
            new ExcelNamedRangeTarget("SUM(Source!$B$2:$B$3)")));
    source.setCell("L1", ExcelCellValue.text("Region"));
    source.setCell("M1", ExcelCellValue.text("Desk"));
    source.setCell("L2", ExcelCellValue.text("North"));
    source.setCell("M2", ExcelCellValue.text("A1"));
    source.setCell("L3", ExcelCellValue.text("South"));
    source.setCell("M3", ExcelCellValue.text("B1"));
    workbook.setTable(
        new ExcelTableDefinition(
            "ReplicaTable",
            "Source",
            "L1:M3",
            false,
            true,
            new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false),
            "replica table",
            true,
            true,
            false,
            "HeaderStyle",
            "DataStyle",
            "TotalsStyle",
            List.of(
                new ExcelTableColumnDefinition(0, "Region", "", "", ""),
                new ExcelTableColumnDefinition(1, "Desk", "", "", "UPPER([@Desk])"))));
    workbook.setSheetProtection("Source", protectionSettings(), "gridgrind-copy");
    return source;
  }

  private static List<ExcelAutofilterFilterColumn> advancedAutofilterCriteria() {
    return List.of(
        new ExcelAutofilterFilterColumn(
            0L, false, new ExcelAutofilterFilterCriterion.Values(List.of("Ada"), true)),
        new ExcelAutofilterFilterColumn(
            1L,
            true,
            new ExcelAutofilterFilterCriterion.Custom(
                true,
                List.of(
                    new ExcelAutofilterFilterCriterion.CustomCondition("greaterThan", "1"),
                    new ExcelAutofilterFilterCriterion.CustomCondition("equal", "Queue")))),
        new ExcelAutofilterFilterColumn(
            2L, true, new ExcelAutofilterFilterCriterion.Dynamic("today", 1.0d, 2.0d)),
        new ExcelAutofilterFilterColumn(
            3L, true, new ExcelAutofilterFilterCriterion.Top10(10, false, true)),
        new ExcelAutofilterFilterColumn(
            4L, true, new ExcelAutofilterFilterCriterion.Color(false, new ExcelColor("#AABBCC"))),
        new ExcelAutofilterFilterColumn(
            5L, true, new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", 2)));
  }

  private static ExcelAutofilterSortState advancedAutofilterSortState() {
    return new ExcelAutofilterSortState(
        "A1:F3",
        true,
        true,
        "stroke",
        List.of(
            new ExcelAutofilterSortCondition(
                "A2:A3", true, "cellColor", new ExcelColor("#102030"), null),
            new ExcelAutofilterSortCondition("B2:B3", false, "icon", null, 4)));
  }

  private static ExcelDifferentialStyleSnapshot supportedStyle() {
    return new ExcelDifferentialStyleSnapshot(
        "0.00", null, null, null, null, null, null, null, null, List.of());
  }

  private static ExcelSheetProtectionSettings protectionSettings() {
    return new ExcelSheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }
}
