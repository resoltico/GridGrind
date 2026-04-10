package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for direct conditional-formatting authoring, introspection, and health analysis. */
class ExcelConditionalFormattingControllerTest {
  private final ExcelConditionalFormattingController controller =
      new ExcelConditionalFormattingController();

  @Test
  void authoredBlocksRoundTripAndSelectedClearRenumbersPriorities() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-conditional-formatting-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      ExcelSheet sheet = workbook.sheet("Ops");
      sheet.setConditionalFormatting(primaryBlock("A1:A3"));
      sheet.setConditionalFormatting(secondaryBlock("C1:C3"));

      assertEquals(
          List.of(
              new ExcelConditionalFormattingBlockSnapshot(
                  List.of("A1:A3"),
                  List.of(
                      new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                          1,
                          true,
                          "A1>10",
                          new ExcelDifferentialStyleSnapshot(
                              "0.00",
                              true,
                              false,
                              ExcelFontHeight.fromPoints(BigDecimal.valueOf(12)),
                              "#102030",
                              true,
                              true,
                              "#E0F0AA",
                              expandedBorder(
                                  new ExcelDifferentialBorderSide(
                                      ExcelBorderStyle.THIN, "#405060")),
                              List.of())),
                      new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
                          2,
                          false,
                          ExcelComparisonOperator.BETWEEN,
                          "1",
                          "9",
                          new ExcelDifferentialStyleSnapshot(
                              null, null, true, null, null, null, null, "#AAEECC", null,
                              List.of())))),
              new ExcelConditionalFormattingBlockSnapshot(
                  List.of("C1:C3"),
                  List.of(
                      new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                          3,
                          false,
                          "C1=\"Ready\"",
                          new ExcelDifferentialStyleSnapshot(
                              null, false, null, null, "#223344", null, null, null, null,
                              List.of()))))),
          sheet.conditionalFormatting(new ExcelRangeSelection.All()));

      sheet.clearConditionalFormatting(new ExcelRangeSelection.Selected(List.of("A2")));

      assertEquals(
          List.of(
              new ExcelConditionalFormattingBlockSnapshot(
                  List.of("C1:C3"),
                  List.of(
                      new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                          1,
                          false,
                          "C1=\"Ready\"",
                          new ExcelDifferentialStyleSnapshot(
                              null, false, null, null, "#223344", null, null, null, null,
                              List.of()))))),
          sheet.conditionalFormatting(new ExcelRangeSelection.All()));

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertEquals(
          List.of(
              new ExcelConditionalFormattingBlockSnapshot(
                  List.of("C1:C3"),
                  List.of(
                      new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                          1,
                          false,
                          "C1=\"Ready\"",
                          new ExcelDifferentialStyleSnapshot(
                              null, false, null, null, "#223344", null, null, null, null,
                              List.of()))))),
          reopened.sheet("Ops").conditionalFormatting(new ExcelRangeSelection.All()));
    }
  }

  @Test
  void healthyGreaterThanRuleSupportsSelectionFilteringAndSheetDelegation() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setCell("A1", ExcelCellValue.number(5));
      sheet.setCell("A2", ExcelCellValue.number(9));
      sheet.setConditionalFormatting(
          new ExcelConditionalFormattingBlockDefinition(
              List.of("A1:A2"),
              List.of(
                  new ExcelConditionalFormattingRule.CellValueRule(
                      ExcelComparisonOperator.GREATER_THAN,
                      "1",
                      null,
                      false,
                      new ExcelDifferentialStyle(
                          "0.00", null, null, null, null, null, null, null, null)))));

      assertEquals(1, sheet.conditionalFormattingBlockCount());
      assertEquals(
          List.of(),
          sheet.conditionalFormatting(new ExcelRangeSelection.Selected(List.of("C1:C2"))));

      List<ExcelConditionalFormattingBlockSnapshot> selected =
          sheet.conditionalFormatting(new ExcelRangeSelection.Selected(List.of("A2")));
      assertEquals(1, selected.size());
      ExcelConditionalFormattingRuleSnapshot.CellValueRule rule =
          assertInstanceOf(
              ExcelConditionalFormattingRuleSnapshot.CellValueRule.class,
              selected.getFirst().rules().getFirst());
      assertEquals(ExcelComparisonOperator.GREATER_THAN, rule.operator());
      assertEquals("1", rule.formula1());
      assertNull(rule.formula2());
      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet.xssfSheet()).stream()
              .noneMatch(
                  finding ->
                      finding.code()
                          == WorkbookAnalysis.AnalysisFindingCode
                              .CONDITIONAL_FORMATTING_PRIORITY_COLLISION));

      sheet.clearConditionalFormatting(new ExcelRangeSelection.All());
      assertEquals(0, sheet.conditionalFormattingBlockCount());
      assertEquals(List.of(), sheet.conditionalFormatting(new ExcelRangeSelection.All()));
    }
  }

  @Test
  void readsAdvancedAndUnsupportedLoadedRuleFamilies() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      var formatting = sheet.getSheetConditionalFormatting();

      XSSFConditionalFormattingRule colorScaleRule =
          formatting.createConditionalFormattingColorScaleRule();
      var colorScale = colorScaleRule.getColorScaleFormatting();
      colorScale.setNumControlPoints(3);
      var colorScaleThresholds = colorScale.getThresholds();
      colorScaleThresholds[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
      colorScaleThresholds[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
      colorScaleThresholds[1].setValue(50d);
      colorScaleThresholds[2].setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
      colorScale.setThresholds(colorScaleThresholds);
      colorScale.setColors(
          new XSSFColor[] {
            new XSSFColor(new byte[] {(byte) 0xAA, 0x00, 0x00}),
            new XSSFColor(new byte[] {(byte) 0xFF, (byte) 0xDD, 0x55}),
            new XSSFColor(new byte[] {0x11, (byte) 0xCC, 0x66})
          });
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("A1:A3")}, colorScaleRule);

      XSSFConditionalFormattingRule dataBarRule =
          formatting.createConditionalFormattingRule(new XSSFColor(new byte[] {0x12, 0x34, 0x56}));
      var dataBar = dataBarRule.getDataBarFormatting();
      dataBar.setIconOnly(true);
      dataBar.setLeftToRight(false);
      dataBar.setWidthMin(10);
      dataBar.setWidthMax(90);
      dataBar.getMinThreshold().setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
      dataBar.getMaxThreshold().setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("B1:B3")}, dataBarRule);

      XSSFConditionalFormattingRule iconSetRule =
          formatting.createConditionalFormattingRule(
              IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS);
      var iconSet = iconSetRule.getMultiStateFormatting();
      iconSet.setIconOnly(true);
      iconSet.setReversed(true);
      var iconThresholds = iconSet.getThresholds();
      iconThresholds[0].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
      iconThresholds[0].setValue(0d);
      iconThresholds[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
      iconThresholds[1].setValue(33d);
      iconThresholds[2].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
      iconThresholds[2].setValue(67d);
      iconSet.setThresholds(iconThresholds);
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("C1:C3")}, iconSetRule);

      XSSFConditionalFormattingRule unsupportedRule =
          formatting.createConditionalFormattingRule("D1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("D1:D3")}, unsupportedRule);
      var unsupportedCtRule = ctBlock(sheet, "D1:D3").getCfRuleArray(0);
      while (unsupportedCtRule.sizeOfFormulaArray() > 0) {
        unsupportedCtRule.removeFormula(0);
      }

      List<ExcelConditionalFormattingBlockSnapshot> blocks =
          controller.conditionalFormatting(sheet, new ExcelRangeSelection.All());

      assertEquals(4, blocks.size());
      assertInstanceOf(
          ExcelConditionalFormattingRuleSnapshot.ColorScaleRule.class,
          blocks.get(0).rules().getFirst());
      assertInstanceOf(
          ExcelConditionalFormattingRuleSnapshot.DataBarRule.class,
          blocks.get(1).rules().getFirst());
      assertInstanceOf(
          ExcelConditionalFormattingRuleSnapshot.IconSetRule.class,
          blocks.get(2).rules().getFirst());
      ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupported =
          assertInstanceOf(
              ExcelConditionalFormattingRuleSnapshot.UnsupportedRule.class,
              blocks.get(3).rules().getFirst());
      assertEquals("FORMULA", unsupported.kind());
      assertEquals("Formula rule is missing formula text.", unsupported.detail());

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.conditionalFormattingHealthFindings("Ops", sheet);
      assertEquals(
          List.of(WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_UNSUPPORTED_RULE),
          findings.stream().map(WorkbookAnalysis.AnalysisFinding::code).distinct().toList());
    }
  }

  @Test
  void healthFindingsDetectBrokenFormulaEmptyRangeAndPriorityCollision() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setConditionalFormatting(sheet, primaryBlock("A1:A3"));
      controller.setConditionalFormatting(sheet, secondaryBlock("B1:B3"));

      var firstRule = ctBlock(sheet, "A1:A3").getCfRuleArray(0);
      firstRule.setFormulaArray(0, "#REF!");
      var secondBlock = ctBlock(sheet, "B1:B3");
      secondBlock.getCfRuleArray(0).setPriority(firstRule.getPriority());
      secondBlock.setSqref(List.of());

      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .distinct()
              .toList();

      assertEquals(
          List.of(
              WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
              WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE,
              WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION),
          codes);
    }
  }

  @Test
  void locationEvidenceFormatsEveryAnalysisLocationFamily() {
    assertEquals(
        "WORKBOOK",
        ExcelConditionalFormattingController.locationEvidence(
            new WorkbookAnalysis.AnalysisLocation.Workbook()));
    assertEquals(
        "Ops",
        ExcelConditionalFormattingController.locationEvidence(
            new WorkbookAnalysis.AnalysisLocation.Sheet("Ops")));
    assertEquals(
        "Ops!A1",
        ExcelConditionalFormattingController.locationEvidence(
            new WorkbookAnalysis.AnalysisLocation.Cell("Ops", "A1")));
    assertEquals(
        "Ops!A1:B2",
        ExcelConditionalFormattingController.locationEvidence(
            new WorkbookAnalysis.AnalysisLocation.Range("Ops", "A1:B2")));
    assertEquals(
        "BudgetTotal@WorkbookScope[]",
        ExcelConditionalFormattingController.locationEvidence(
            new WorkbookAnalysis.AnalysisLocation.NamedRange(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
  }

  @Test
  void toSnapshotDegradesUnsupportedAdvancedPayloads() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      var formatting = sheet.getSheetConditionalFormatting();

      XSSFConditionalFormattingRule colorScaleRule =
          formatting.createConditionalFormattingColorScaleRule();
      var colorScale = colorScaleRule.getColorScaleFormatting();
      colorScale.setNumControlPoints(2);
      colorScale.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
      colorScale.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
      colorScale.setColors(
          new XSSFColor[] {
            new XSSFColor(new byte[] {(byte) 0xAA, 0x22, 0x11}),
            new XSSFColor(new byte[] {0x11, (byte) 0xCC, 0x66})
          });
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("A1:A3")}, colorScaleRule);
      var colorScaleCtRule = ctBlock(sheet, "A1:A3").getCfRuleArray(0);
      colorScaleCtRule.getColorScale().getColorArray(0).setRgb(new byte[] {0x01, 0x02});

      XSSFConditionalFormattingRule dataBarRule =
          formatting.createConditionalFormattingRule(new XSSFColor(new byte[] {0x12, 0x34, 0x56}));
      var dataBar = dataBarRule.getDataBarFormatting();
      dataBar.getMinThreshold().setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
      dataBar.getMaxThreshold().setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("B1:B3")}, dataBarRule);
      var dataBarCtRule = ctBlock(sheet, "B1:B3").getCfRuleArray(0);
      dataBarCtRule.getDataBar().getColor().setRgb(new byte[] {0x01, 0x02});

      XSSFConditionalFormattingRule iconSetRule =
          formatting.createConditionalFormattingRule(
              IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS);
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("C1:C3")}, iconSetRule);
      var iconSetCtRule = ctBlock(sheet, "C1:C3").getCfRuleArray(0);
      iconSetCtRule.unsetIconSet();

      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(0).getRule(0), colorScaleCtRule),
          "COLOR_SCALE",
          "unsupported threshold or color payload");
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(1).getRule(0), dataBarCtRule),
          "DATA_BAR",
          "unsupported threshold or color payload");
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(2).getRule(0), iconSetCtRule),
          "ICON_SET",
          "missing icon-set payload");
    }
  }

  @Test
  void helperMethodsExposeStableLabelsAndFormulaValidation() throws Exception {
    ExcelDifferentialStyleSnapshot style =
        new ExcelDifferentialStyleSnapshot(
            "0.00", null, null, null, null, null, null, null, null, List.of());

    assertEquals(
        "FORMULA_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ExcelConditionalFormattingRuleSnapshot.FormulaRule(1, false, "A1>0", style)));
    assertEquals(
        "CELL_VALUE_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
                2, false, ExcelComparisonOperator.GREATER_THAN, "1", null, style)));
    assertEquals(
        "COLOR_SCALE_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
                3,
                false,
                List.of(
                    new ExcelConditionalFormattingThresholdSnapshot(
                        ExcelConditionalFormattingThresholdType.MIN, null, 0.0d)),
                List.of("#102030"))));
    assertEquals(
        "DATA_BAR_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                4,
                false,
                "#223344",
                false,
                true,
                0,
                100,
                new ExcelConditionalFormattingThresholdSnapshot(
                    ExcelConditionalFormattingThresholdType.MIN, null, 0.0d),
                new ExcelConditionalFormattingThresholdSnapshot(
                    ExcelConditionalFormattingThresholdType.MAX, null, 1.0d))));
    assertEquals(
        "ICON_SET_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
                5,
                false,
                ExcelConditionalFormattingIconSet.GYR_3_ARROW,
                false,
                false,
                List.of(
                    new ExcelConditionalFormattingThresholdSnapshot(
                        ExcelConditionalFormattingThresholdType.MIN, null, 0.0d)))));
    assertEquals(
        "UNSUPPORTED_RULE(DATA_BAR)",
        ExcelConditionalFormattingController.ruleLabel(
            new ExcelConditionalFormattingRuleSnapshot.UnsupportedRule(
                6, false, "data_bar", "detail")));
    var explicitUnsupportedType =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule.Factory.newInstance();
    explicitUnsupportedType.setType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType.ABOVE_AVERAGE);
    assertEquals(
        "ABOVE_AVERAGE",
        ExcelConditionalFormattingController.unsupportedKind(
            org.apache.poi.ss.usermodel.ConditionType.FORMULA, explicitUnsupportedType));
    assertEquals("TOP_10", ExcelConditionalFormattingController.normalizedUnsupportedKind("top10"));
    assertEquals(
        "ABOVE_AVERAGE",
        ExcelConditionalFormattingController.normalizedUnsupportedKind("aboveAverage"));
    assertEquals(
        "FORMULA",
        ExcelConditionalFormattingController.unsupportedKind(
            org.apache.poi.ss.usermodel.ConditionType.FORMULA,
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule.Factory.newInstance()));
    assertEquals(
        "CELL_VALUE_IS",
        ExcelConditionalFormattingController.unsupportedKind(
            org.apache.poi.ss.usermodel.ConditionType.CELL_VALUE_IS,
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule.Factory.newInstance()));
    assertEquals(
        "COLOR_SCALE",
        ExcelConditionalFormattingController.unsupportedKind(
            org.apache.poi.ss.usermodel.ConditionType.COLOR_SCALE,
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule.Factory.newInstance()));
    assertEquals(
        "DATA_BAR",
        ExcelConditionalFormattingController.unsupportedKind(
            org.apache.poi.ss.usermodel.ConditionType.DATA_BAR,
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule.Factory.newInstance()));
    assertEquals(
        "ICON_SET",
        ExcelConditionalFormattingController.unsupportedKind(
            org.apache.poi.ss.usermodel.ConditionType.ICON_SET,
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule.Factory.newInstance()));
    assertFalse(
        ExcelConditionalFormattingController.unsupportedKind(
                org.apache.poi.ss.usermodel.ConditionType.FILTER,
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule.Factory.newInstance())
            .isBlank());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(workbook);
      int sheetIndex = workbook.getSheetIndex(sheet);

      assertFalse(
          ExcelConditionalFormattingController.isBrokenFormula(
              evaluationWorkbook, sheetIndex, "A1>0"));
      assertTrue(
          ExcelConditionalFormattingController.isBrokenFormula(
              evaluationWorkbook, sheetIndex, "#REF!"));
      assertTrue(
          ExcelConditionalFormattingController.isBrokenFormula(
              evaluationWorkbook, sheetIndex, "SUM("));
    }
  }

  @Test
  void healthFindingsUseSingleOperandEvidenceForBrokenCellValueRules() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setConditionalFormatting(
          sheet,
          new ExcelConditionalFormattingBlockDefinition(
              List.of("A1:A2"),
              List.of(
                  new ExcelConditionalFormattingRule.CellValueRule(
                      ExcelComparisonOperator.GREATER_THAN,
                      "#REF!",
                      null,
                      false,
                      new ExcelDifferentialStyle(
                          "0.00", null, null, null, null, null, null, null, null)))));

      WorkbookAnalysis.AnalysisFinding finding =
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .filter(
                  candidate ->
                      candidate.code()
                          == WorkbookAnalysis.AnalysisFindingCode
                              .CONDITIONAL_FORMATTING_BROKEN_FORMULA)
              .findFirst()
              .orElseThrow();

      assertEquals(List.of("#REF!"), finding.evidence());
    }
  }

  @Test
  void healthFindingsUseBothOperandsWhenSecondCellValueFormulaIsBroken() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setConditionalFormatting(
          sheet,
          new ExcelConditionalFormattingBlockDefinition(
              List.of("A1:A2"),
              List.of(
                  new ExcelConditionalFormattingRule.CellValueRule(
                      ExcelComparisonOperator.BETWEEN,
                      "1",
                      "#REF!",
                      false,
                      new ExcelDifferentialStyle(
                          "0.00", null, null, null, null, null, null, null, null)))));

      WorkbookAnalysis.AnalysisFinding finding =
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .filter(
                  candidate ->
                      candidate.code()
                          == WorkbookAnalysis.AnalysisFindingCode
                              .CONDITIONAL_FORMATTING_BROKEN_FORMULA)
              .findFirst()
              .orElseThrow();

      assertEquals(List.of("1", "#REF!"), finding.evidence());
    }
  }

  @Test
  void healthFindingsDoNotFlagHealthyFormulaRules() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      controller.setConditionalFormatting(
          sheet,
          new ExcelConditionalFormattingBlockDefinition(
              List.of("A1:A3"),
              List.of(
                  new ExcelConditionalFormattingRule.FormulaRule(
                      "A1>0",
                      false,
                      new ExcelDifferentialStyle(
                          "0.00", null, null, null, null, null, null, null, null)))));

      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .noneMatch(
                  finding ->
                      finding.code()
                          == WorkbookAnalysis.AnalysisFindingCode
                              .CONDITIONAL_FORMATTING_BROKEN_FORMULA));
    }
  }

  @Test
  void healthFindingsFlagMalformedRangesAndNonPositivePriorities() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      var malformedBlock = sheet.getCTWorksheet().addNewConditionalFormatting();
      malformedBlock.setSqref(List.of("not-a-range"));
      var malformedRule = malformedBlock.addNewCfRule();
      malformedRule.setType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType.EXPRESSION);
      malformedRule.setPriority(0);
      malformedRule.addFormula("A1>0");

      assertEquals(
          List.of(),
          controller.conditionalFormatting(
              sheet, new ExcelRangeSelection.Selected(List.of("A1:A3"))));
      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE,
                      WorkbookAnalysis.AnalysisFindingCode
                          .CONDITIONAL_FORMATTING_PRIORITY_COLLISION)));
    }
  }

  @Test
  void toSnapshotCoversMissingPayloadUnsupportedOperatorAndFallbackFamilies() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      var formatting = sheet.getSheetConditionalFormatting();

      XSSFConditionalFormattingRule missingColorScale =
          formatting.createConditionalFormattingColorScaleRule();
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("A1:A3")}, missingColorScale);
      var missingColorScaleCtRule = ctBlock(sheet, "A1:A3").getCfRuleArray(0);
      missingColorScaleCtRule.unsetColorScale();
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(0).getRule(0), missingColorScaleCtRule),
          "COLOR_SCALE",
          "missing color-scale payload");

      XSSFConditionalFormattingRule missingDataBar =
          formatting.createConditionalFormattingRule(new XSSFColor(new byte[] {0x12, 0x34, 0x56}));
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("B1:B3")}, missingDataBar);
      var missingDataBarCtRule = ctBlock(sheet, "B1:B3").getCfRuleArray(0);
      missingDataBarCtRule.unsetDataBar();
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(1).getRule(0), missingDataBarCtRule),
          "DATA_BAR",
          "missing data-bar payload");

      XSSFConditionalFormattingRule blankCellValue =
          formatting.createConditionalFormattingRule(
              org.apache.poi.ss.usermodel.ComparisonOperator.GT, "1");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("C1:C3")}, blankCellValue);
      var blankCellValueCtRule = ctBlock(sheet, "C1:C3").getCfRuleArray(0);
      while (blankCellValueCtRule.sizeOfFormulaArray() > 0) {
        blankCellValueCtRule.removeFormula(0);
      }
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(2).getRule(0), blankCellValueCtRule),
          "CELL_VALUE_IS",
          "missing formula1");

      XSSFConditionalFormattingRule unsupportedOperator =
          formatting.createConditionalFormattingRule(
              org.apache.poi.ss.usermodel.ComparisonOperator.GT, "1");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("D1:D3")}, unsupportedOperator);
      var unsupportedOperatorCtRule = ctBlock(sheet, "D1:D3").getCfRuleArray(0);
      unsupportedOperatorCtRule.unsetOperator();
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet,
              formatting.getConditionalFormattingAt(3).getRule(0),
              unsupportedOperatorCtRule),
          "CELL_VALUE_IS",
          "unsupported comparison payload");

      XSSFConditionalFormattingRule unsupportedFamily =
          formatting.createConditionalFormattingRule("E1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("E1:E3")}, unsupportedFamily);
      var unsupportedFamilyCtRule = ctBlock(sheet, "E1:E3").getCfRuleArray(0);
      unsupportedFamilyCtRule.setType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType.TOP_10);
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(4).getRule(0), unsupportedFamilyCtRule),
          "TOP_10",
          "not modeled");

      XSSFConditionalFormattingRule iconSetCatchRule =
          formatting.createConditionalFormattingRule(IconMultiStateFormatting.IconSet.GYR_3_ARROW);
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("F1:F3")}, iconSetCatchRule);
      var iconSetCatchCtRule = ctBlock(sheet, "F1:F3").getCfRuleArray(0);
      iconSetCatchCtRule
          .getIconSet()
          .getCfvoArray(0)
          .setType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfvoType.FORMULA);
      iconSetCatchCtRule.getIconSet().getCfvoArray(0).setVal(" ");
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(5).getRule(0), iconSetCatchCtRule),
          "ICON_SET",
          "unsupported threshold or icon-set payload");
    }
  }

  @Test
  void excelSheetDelegatesConditionalFormattingHealthAnalysis() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setCell("A1", ExcelCellValue.number(1));
      sheet.setConditionalFormatting(
          new ExcelConditionalFormattingBlockDefinition(
              List.of("A1:A2"),
              List.of(
                  new ExcelConditionalFormattingRule.FormulaRule(
                      "#REF!",
                      false,
                      new ExcelDifferentialStyle(
                          "0.00", null, null, null, null, null, null, null, null)))));

      assertEquals(
          List.of(WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA),
          sheet.conditionalFormattingHealthFindings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .distinct()
              .toList());
    }
  }

  private static ExcelConditionalFormattingBlockDefinition primaryBlock(String range) {
    return new ExcelConditionalFormattingBlockDefinition(
        List.of(range),
        List.of(
            new ExcelConditionalFormattingRule.FormulaRule(
                "A1>10",
                true,
                new ExcelDifferentialStyle(
                    "0.00",
                    true,
                    false,
                    ExcelFontHeight.fromPoints(BigDecimal.valueOf(12)),
                    "#102030",
                    true,
                    true,
                    "#E0F0AA",
                    new ExcelDifferentialBorder(
                        new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#405060"),
                        null,
                        null,
                        null,
                        null))),
            new ExcelConditionalFormattingRule.CellValueRule(
                ExcelComparisonOperator.BETWEEN,
                "1",
                "9",
                false,
                new ExcelDifferentialStyle(
                    null, null, true, null, null, null, null, "#AAEECC", null))));
  }

  private static ExcelConditionalFormattingBlockDefinition secondaryBlock(String range) {
    return new ExcelConditionalFormattingBlockDefinition(
        List.of(range),
        List.of(
            new ExcelConditionalFormattingRule.FormulaRule(
                range.startsWith("B") ? "B1>5" : "C1=\"Ready\"",
                false,
                new ExcelDifferentialStyle(
                    null, false, null, null, "#223344", null, null, null, null))));
  }

  private static ExcelDifferentialBorder expandedBorder(ExcelDifferentialBorderSide side) {
    return new ExcelDifferentialBorder(null, side, side, side, side);
  }

  private static void seedConditionalFormattingValues(XSSFSheet sheet) {
    for (int rowIndex = 0; rowIndex < 3; rowIndex++) {
      sheet.createRow(rowIndex);
      sheet.getRow(rowIndex).createCell(0).setCellValue(rowIndex + 1);
      sheet.getRow(rowIndex).createCell(1).setCellValue(rowIndex + 3);
      sheet.getRow(rowIndex).createCell(2).setCellValue(rowIndex + 5);
      sheet.getRow(rowIndex).createCell(3).setCellValue(rowIndex + 7);
    }
  }

  private static org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting
      ctBlock(XSSFSheet sheet, String range) {
    for (var candidate : sheet.getCTWorksheet().getConditionalFormattingArray()) {
      if (List.of(range).equals(ExcelSqrefSupport.normalizedSqref(candidate.getSqref()))) {
        return candidate;
      }
    }
    throw new IllegalStateException("conditional-formatting block not found for " + range);
  }

  private static void assertUnsupportedRule(
      ExcelConditionalFormattingRuleSnapshot snapshot, String kind, String detailFragment) {
    ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupported =
        assertInstanceOf(ExcelConditionalFormattingRuleSnapshot.UnsupportedRule.class, snapshot);
    assertEquals(kind, unsupported.kind());
    assertTrue(unsupported.detail().contains(detailFragment), unsupported.detail());
  }
}
