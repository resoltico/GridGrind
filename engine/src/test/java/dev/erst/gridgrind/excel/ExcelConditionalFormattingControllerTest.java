package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
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
  @Test
  void invalidTargetRangeListsPreserveExistingRules() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      var controller = new ExcelConditionalFormattingController();
      var range = new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:A3");
      var formatting =
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.FormulaRule("A1>0", true, Optional.empty())));
      controller.setConditionalFormatting(sheet, List.of(range), formatting);
      assertThrows(
          IllegalArgumentException.class,
          () -> controller.setConditionalFormatting(sheet, List.of(), formatting));
      assertThrows(
          IllegalArgumentException.class,
          () -> controller.setConditionalFormatting(sheet, List.of(range, range), formatting));
      assertEquals(1, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
      assertEquals(
          "A1>0",
          sheet
              .getSheetConditionalFormatting()
              .getConditionalFormattingAt(0)
              .getRule(0)
              .getFormula1());
    }
  }

  @Test
  void normalizesAuthoredFormulasAtPoiBoundary() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      var controller = new ExcelConditionalFormattingController();
      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.FormulaRule(
                      " =A1>0 ", true, Optional.empty()))));
      assertEquals(
          "A1>0",
          sheet
              .getSheetConditionalFormatting()
              .getConditionalFormattingAt(0)
              .getRule(0)
              .getFormula1());
    }
  }

  @Test
  void invalidReplacementFormulaPreservesExistingConditionalFormatting() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      var controller = new ExcelConditionalFormattingController();
      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.FormulaRule("A1>0", true, Optional.empty()))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setConditionalFormatting(
                  sheet,
                  List.of("A1:A3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                      .toList(),
                  new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                      List.of(
                          new ConditionalFormattingRuleSpec.FormulaRule(
                              " = ", true, Optional.empty())))));
      assertEquals(1, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
      assertEquals(
          "A1>0",
          sheet
              .getSheetConditionalFormatting()
              .getConditionalFormattingAt(0)
              .getRule(0)
              .getFormula1());
    }
  }

  private final ExcelConditionalFormattingController controller =
      new ExcelConditionalFormattingController();

  @Test
  void authoredBlocksRoundTripAndSelectedClearRenumbersPriorities() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-conditional-formatting-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      ExcelSheet sheet = workbook.sheet("Ops");
      sheet
          .metadata()
          .setConditionalFormatting(
              List.of(new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:A3")), primaryBlock());
      sheet
          .metadata()
          .setConditionalFormatting(
              List.of(new dev.erst.gridgrind.core.vocabulary.RangeRef("C1:C3")),
              secondaryBlock("C1:C3"));

      assertEquals(
          List.of(
              new ConditionalFormattingEntryReport(
                  List.of("A1:A3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.FormulaRule(
                          1,
                          true,
                          "A1>10",
                          Optional.ofNullable(
                              new DifferentialStyleReport(
                                  "0.00",
                                  true,
                                  false,
                                  FontHeightReport.fromPoints(BigDecimal.valueOf(12)),
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                                          color("#102030"))),
                                  true,
                                  true,
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                                          color("#E0F0AA"))),
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.border(
                                          expandedBorder(
                                              borderSide(ExcelBorderStyle.THIN, "#405060")))),
                                  List.of()))),
                      new ConditionalFormattingRuleReport.CellValueRule(
                          2,
                          false,
                          ExcelComparisonOperator.BETWEEN,
                          "1",
                          Optional.ofNullable("9"),
                          Optional.ofNullable(
                              new DifferentialStyleReport(
                                  null,
                                  null,
                                  true,
                                  null,
                                  Optional.empty(),
                                  null,
                                  null,
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                                          color("#AAEECC"))),
                                  Optional.empty(),
                                  List.of()))))),
              new ConditionalFormattingEntryReport(
                  List.of("C1:C3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.FormulaRule(
                          3,
                          false,
                          "C1=\"Ready\"",
                          Optional.ofNullable(
                              new DifferentialStyleReport(
                                  null,
                                  false,
                                  null,
                                  null,
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                                          color("#223344"))),
                                  null,
                                  null,
                                  Optional.empty(),
                                  Optional.empty(),
                                  List.of())))))),
          sheet.metadata().conditionalFormatting(new ExcelRangeSelection.All()));

      sheet.metadata().clearConditionalFormatting(new ExcelRangeSelection.Selected(List.of("A2")));

      assertEquals(
          List.of(
              new ConditionalFormattingEntryReport(
                  List.of("C1:C3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.FormulaRule(
                          1,
                          false,
                          "C1=\"Ready\"",
                          Optional.ofNullable(
                              new DifferentialStyleReport(
                                  null,
                                  false,
                                  null,
                                  null,
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                                          color("#223344"))),
                                  null,
                                  null,
                                  Optional.empty(),
                                  Optional.empty(),
                                  List.of())))))),
          sheet.metadata().conditionalFormatting(new ExcelRangeSelection.All()));

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(
          List.of(
              new ConditionalFormattingEntryReport(
                  List.of("C1:C3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.FormulaRule(
                          1,
                          false,
                          "C1=\"Ready\"",
                          Optional.ofNullable(
                              new DifferentialStyleReport(
                                  null,
                                  false,
                                  null,
                                  null,
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                                          color("#223344"))),
                                  null,
                                  null,
                                  Optional.empty(),
                                  Optional.empty(),
                                  List.of())))))),
          reopened.sheet("Ops").metadata().conditionalFormatting(new ExcelRangeSelection.All()));
    }
  }

  @Test
  void authoredBarriersCanPersistWithoutStyles() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);

      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.FormulaRule("A1>0", true, Optional.empty()),
                  new ConditionalFormattingRuleSpec.CellValueRule(
                      ExcelComparisonOperator.GREATER_THAN,
                      "1",
                      Optional.empty(),
                      true,
                      Optional.empty()),
                  new ConditionalFormattingRuleSpec.Top10Rule(
                      true, 10, false, false, Optional.empty()))));

      assertEquals(
          List.of(
              new ConditionalFormattingEntryReport(
                  List.of("A1:A3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.FormulaRule(
                          1, true, "A1>0", Optional.empty()),
                      new ConditionalFormattingRuleReport.CellValueRule(
                          2,
                          true,
                          ExcelComparisonOperator.GREATER_THAN,
                          "1",
                          Optional.empty(),
                          Optional.empty()),
                      new ConditionalFormattingRuleReport.Top10Rule(
                          3, true, 10, false, false, Optional.empty())))),
          controller.conditionalFormatting(sheet, new ExcelRangeSelection.All()));
    }
  }

  @Test
  void authoredAdvancedThresholdRulesSupportMinMaxWithoutValues() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);

      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.ColorScaleRule(
                      false,
                      List.of(
                          new ConditionalFormattingThresholdSpec.Min(),
                          new ConditionalFormattingThresholdSpec.Percentile(50.0d),
                          new ConditionalFormattingThresholdSpec.Max()),
                      List.of(
                          ColorSpec.rgb(new RgbHex("#AA2211")),
                          ColorSpec.rgb(new RgbHex("#FFDD55")),
                          ColorSpec.rgb(new RgbHex("#11CC66")))),
                  new ConditionalFormattingRuleSpec.DataBarRule(
                      false,
                      ColorSpec.rgb(new RgbHex("#123456")),
                      true,
                      10,
                      90,
                      new ConditionalFormattingThresholdSpec.Min(),
                      new ConditionalFormattingThresholdSpec.Max()))));

      assertEquals(
          List.of(
              new ConditionalFormattingEntryReport(
                  List.of("A1:A3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.ColorScaleRule(
                          1,
                          false,
                          List.of(
                              new ConditionalFormattingThresholdReport(
                                  ExcelConditionalFormattingThresholdType.MIN, null, null),
                              new ConditionalFormattingThresholdReport(
                                  ExcelConditionalFormattingThresholdType.PERCENTILE, null, 50.0d),
                              new ConditionalFormattingThresholdReport(
                                  ExcelConditionalFormattingThresholdType.MAX, null, null)),
                          List.of("#AA2211", "#FFDD55", "#11CC66").stream()
                              .map(dev.erst.gridgrind.core.vocabulary.RgbHex::new)
                              .toList()),
                      new ConditionalFormattingRuleReport.DataBarRule(
                          2,
                          false,
                          Optional.ofNullable("#123456")
                              .map(dev.erst.gridgrind.core.vocabulary.RgbHex::new),
                          true,
                          10,
                          90,
                          new ConditionalFormattingThresholdReport(
                              ExcelConditionalFormattingThresholdType.MIN, null, null),
                          new ConditionalFormattingThresholdReport(
                              ExcelConditionalFormattingThresholdType.MAX, null, null))))),
          controller.conditionalFormatting(sheet, new ExcelRangeSelection.All()));
    }
  }

  @Test
  void authoredIconSetRulesRoundTripFormulaAndValueThresholds() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);

      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.IconSetRule(
                      true,
                      ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                      true,
                      false,
                      List.of(
                          new ConditionalFormattingThresholdSpec.Formula("A1*2"),
                          new ConditionalFormattingThresholdSpec.Numeric(40.0d),
                          new ConditionalFormattingThresholdSpec.Max())))));

      assertEquals(
          List.of(
              new ConditionalFormattingEntryReport(
                  List.of("A1:A3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.IconSetRule(
                          1,
                          true,
                          ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                          true,
                          false,
                          List.of(
                              new ConditionalFormattingThresholdReport(
                                  ExcelConditionalFormattingThresholdType.FORMULA, "A1*2", null),
                              new ConditionalFormattingThresholdReport(
                                  ExcelConditionalFormattingThresholdType.NUMBER, null, 40.0d),
                              new ConditionalFormattingThresholdReport(
                                  ExcelConditionalFormattingThresholdType.MAX, null, null)))))),
          controller.conditionalFormatting(sheet, new ExcelRangeSelection.All()));
    }
  }

  @Test
  void authoredTop10RulesRoundTripStyledPayloads() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);

      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.Top10Rule(
                      true,
                      7,
                      true,
                      true,
                      Optional.of(
                          new DifferentialStylePatch(
                              Optional.empty(),
                              Optional.of(true),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.of(color("#223344")),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))))));

      assertEquals(
          List.of(
              new ConditionalFormattingEntryReport(
                  List.of("A1:A3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  List.of(
                      new ConditionalFormattingRuleReport.Top10Rule(
                          1,
                          true,
                          7,
                          true,
                          true,
                          Optional.ofNullable(
                              new DifferentialStyleReport(
                                  null,
                                  true,
                                  null,
                                  null,
                                  Optional.ofNullable(
                                      dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                                          color("#223344"))),
                                  null,
                                  null,
                                  Optional.empty(),
                                  Optional.empty(),
                                  List.of())))))),
          controller.conditionalFormatting(sheet, new ExcelRangeSelection.All()));
    }
  }

  @Test
  void healthyGreaterThanRuleSupportsSelectionFilteringAndSheetDelegation() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.cells().setCell("A1", new CellValue.NumberValue(5));
      sheet.cells().setCell("A2", new CellValue.NumberValue(9));
      sheet
          .metadata()
          .setConditionalFormatting(
              List.of("A1:A2").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                  .toList(),
              new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                  List.of(
                      new ConditionalFormattingRuleSpec.CellValueRule(
                          ExcelComparisonOperator.GREATER_THAN,
                          "1",
                          Optional.empty(),
                          false,
                          Optional.of(
                              new DifferentialStylePatch(
                                  Optional.of("0.00"),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty()))))));

      assertEquals(1, sheet.diagnostics().conditionalFormattingBlockCount());
      assertEquals(
          List.of(),
          sheet
              .metadata()
              .conditionalFormatting(new ExcelRangeSelection.Selected(List.of("C1:C2"))));

      List<ConditionalFormattingEntryReport> selected =
          sheet.metadata().conditionalFormatting(new ExcelRangeSelection.Selected(List.of("A2")));
      assertEquals(1, selected.size());
      ConditionalFormattingRuleReport.CellValueRule rule =
          assertInstanceOf(
              ConditionalFormattingRuleReport.CellValueRule.class,
              selected.getFirst().rules().getFirst());
      assertEquals(ExcelComparisonOperator.GREATER_THAN, rule.operator());
      assertEquals("1", rule.formula1());
      assertTrue(rule.formula2().isEmpty());
      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet.xssfSheet()).stream()
              .noneMatch(
                  finding ->
                      finding.code()
                          == AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION));

      sheet.metadata().clearConditionalFormatting(new ExcelRangeSelection.All());
      assertEquals(0, sheet.diagnostics().conditionalFormattingBlockCount());
      assertEquals(
          List.of(), sheet.metadata().conditionalFormatting(new ExcelRangeSelection.All()));
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

      List<ConditionalFormattingEntryReport> blocks =
          controller.conditionalFormatting(sheet, new ExcelRangeSelection.All());

      assertEquals(4, blocks.size());
      assertInstanceOf(
          ConditionalFormattingRuleReport.ColorScaleRule.class, blocks.get(0).rules().getFirst());
      assertInstanceOf(
          ConditionalFormattingRuleReport.DataBarRule.class, blocks.get(1).rules().getFirst());
      assertInstanceOf(
          ConditionalFormattingRuleReport.IconSetRule.class, blocks.get(2).rules().getFirst());
      ConditionalFormattingRuleReport.UnsupportedRule unsupported =
          assertInstanceOf(
              ConditionalFormattingRuleReport.UnsupportedRule.class,
              blocks.get(3).rules().getFirst());
      assertEquals("FORMULA", unsupported.kind());
      assertEquals("Formula rule is missing formula text.", unsupported.detail());

      List<AnalysisFindingReport> findings =
          controller.conditionalFormattingHealthFindings("Ops", sheet);
      assertEquals(
          List.of(AnalysisFindingCode.CONDITIONAL_FORMATTING_UNSUPPORTED_RULE),
          findings.stream().map(AnalysisFindingReport::code).distinct().toList());
    }
  }

  @Test
  void healthFindingsDetectBrokenFormulaEmptyRangeAndPriorityCollision() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setConditionalFormatting(
          sheet, List.of(new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:A3")), primaryBlock());
      controller.setConditionalFormatting(
          sheet,
          List.of(new dev.erst.gridgrind.core.vocabulary.RangeRef("B1:B3")),
          secondaryBlock("B1:B3"));

      var firstRule = ctBlock(sheet, "A1:A3").getCfRuleArray(0);
      firstRule.setFormulaArray(0, "#REF!");
      var secondBlock = ctBlock(sheet, "B1:B3");
      secondBlock.getCfRuleArray(0).setPriority(firstRule.getPriority());
      secondBlock.setSqref(List.of());

      List<AnalysisFindingCode> codes =
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .map(AnalysisFindingReport::code)
              .distinct()
              .toList();

      assertEquals(
          List.of(
              AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
              AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE,
              AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION),
          codes);
    }
  }

  @Test
  void locationEvidenceFormatsEveryAnalysisLocationFamily() {
    assertEquals(
        "WORKBOOK",
        ExcelConditionalFormattingController.locationEvidence(
            new AnalysisLocationReport.Workbook()));
    assertEquals(
        "Ops",
        ExcelConditionalFormattingController.locationEvidence(
            new AnalysisLocationReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"))));
    assertEquals(
        "Ops!A1",
        ExcelConditionalFormattingController.locationEvidence(
            new AnalysisLocationReport.Cell(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"), "A1")));
    assertEquals(
        "Ops!A1:B2",
        ExcelConditionalFormattingController.locationEvidence(
            new AnalysisLocationReport.Range(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"), "A1:B2")));
    assertEquals(
        "BudgetTotal@workbook",
        ExcelConditionalFormattingController.locationEvidence(
            new AnalysisLocationReport.NamedRange(
                "BudgetTotal", new NamedRangeScopeReport.Workbook())));
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
  void toSnapshotFallsBackToConditionTypeWhenRawCfTypeIsUnset() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      var formatting = sheet.getSheetConditionalFormatting();
      XSSFConditionalFormattingRule formulaRule =
          formatting.createConditionalFormattingRule("A1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("A1:A3")}, formulaRule);
      var ctRule = ctBlock(sheet, "A1:A3").getCfRuleArray(0);
      ctRule.unsetType();

      ConditionalFormattingRuleReport.UnsupportedRule snapshot =
          assertInstanceOf(
              ConditionalFormattingRuleReport.UnsupportedRule.class,
              ExcelConditionalFormattingController.toSnapshot(
                  sheet, formatting.getConditionalFormattingAt(0).getRule(0), ctRule));

      assertEquals("UNKNOWN", snapshot.kind());
      assertEquals("Rule family metadata is missing or unreadable.", snapshot.detail());
    }
  }

  @Test
  void toSnapshotDegradesUnreadableAndUnsupportedRuleFamiliesInsteadOfCrashing() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      var formatting = sheet.getSheetConditionalFormatting();

      XSSFConditionalFormattingRule unreadableRule =
          formatting.createConditionalFormattingRule("A1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("A1:A3")}, unreadableRule);
      var unreadableCtRule = ctBlock(sheet, "A1:A3").getCfRuleArray(0);
      try (var cursor = unreadableCtRule.newCursor()) {
        cursor.setAttributeText(new javax.xml.namespace.QName("", "type"), "mystery");
      }
      ConditionalFormattingRuleReport.UnsupportedRule unreadableSnapshot =
          assertInstanceOf(
              ConditionalFormattingRuleReport.UnsupportedRule.class,
              ExcelConditionalFormattingController.toSnapshot(
                  sheet, formatting.getConditionalFormattingAt(0).getRule(0), unreadableCtRule));
      assertEquals("UNKNOWN", unreadableSnapshot.kind());
      assertEquals("Rule family metadata is missing or unreadable.", unreadableSnapshot.detail());

      XSSFConditionalFormattingRule unsupportedFamilyRule =
          formatting.createConditionalFormattingRule("B1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("B1:B3")}, unsupportedFamilyRule);
      var unsupportedFamilyCtRule = ctBlock(sheet, "B1:B3").getCfRuleArray(0);
      unsupportedFamilyCtRule.setType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType.ABOVE_AVERAGE);
      ConditionalFormattingRuleReport.UnsupportedRule unsupportedFamilySnapshot =
          assertInstanceOf(
              ConditionalFormattingRuleReport.UnsupportedRule.class,
              ExcelConditionalFormattingController.toSnapshot(
                  sheet,
                  formatting.getConditionalFormattingAt(1).getRule(0),
                  unsupportedFamilyCtRule));
      assertEquals("ABOVE_AVERAGE", unsupportedFamilySnapshot.kind());
      assertEquals("Rule family is not modeled by GridGrind.", unsupportedFamilySnapshot.detail());
    }
  }

  @Test
  void helperMethodsExposeStableLabelsAndFormulaValidation() throws Exception {
    DifferentialStyleReport style =
        new DifferentialStyleReport(
            "0.00",
            null,
            null,
            null,
            Optional.empty(),
            null,
            null,
            Optional.empty(),
            Optional.empty(),
            List.of());

    assertEquals(
        "FORMULA_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ConditionalFormattingRuleReport.FormulaRule(
                1, false, "A1>0", Optional.ofNullable(style))));
    assertEquals(
        "CELL_VALUE_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ConditionalFormattingRuleReport.CellValueRule(
                2,
                false,
                ExcelComparisonOperator.GREATER_THAN,
                "1",
                Optional.empty(),
                Optional.ofNullable(style))));
    assertEquals(
        "COLOR_SCALE_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ConditionalFormattingRuleReport.ColorScaleRule(
                3,
                false,
                List.of(
                    new ConditionalFormattingThresholdReport(
                        ExcelConditionalFormattingThresholdType.MIN, null, 0.0d)),
                List.of("#102030").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RgbHex::new)
                    .toList())));
    assertEquals(
        "DATA_BAR_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ConditionalFormattingRuleReport.DataBarRule(
                4,
                false,
                Optional.ofNullable("#223344").map(dev.erst.gridgrind.core.vocabulary.RgbHex::new),
                false,
                0,
                100,
                new ConditionalFormattingThresholdReport(
                    ExcelConditionalFormattingThresholdType.MIN, null, 0.0d),
                new ConditionalFormattingThresholdReport(
                    ExcelConditionalFormattingThresholdType.MAX, null, 1.0d))));
    assertEquals(
        "ICON_SET_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ConditionalFormattingRuleReport.IconSetRule(
                5,
                false,
                ExcelConditionalFormattingIconSet.GYR_3_ARROW,
                false,
                false,
                List.of(
                    new ConditionalFormattingThresholdReport(
                        ExcelConditionalFormattingThresholdType.MIN, null, 0.0d)))));
    assertEquals(
        "TOP10_RULE",
        ExcelConditionalFormattingController.ruleLabel(
            new ConditionalFormattingRuleReport.Top10Rule(
                6, true, 7, true, true, Optional.ofNullable(style))));
    assertEquals(
        "UNSUPPORTED_RULE(DATA_BAR)",
        ExcelConditionalFormattingController.ruleLabel(
            new ConditionalFormattingRuleReport.UnsupportedRule(7, false, "data_bar", "detail")));
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
          List.of("A1:A2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.CellValueRule(
                      ExcelComparisonOperator.GREATER_THAN,
                      "#REF!",
                      Optional.empty(),
                      false,
                      Optional.of(
                          new DifferentialStylePatch(
                              Optional.of("0.00"),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))))));

      AnalysisFindingReport finding =
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .filter(
                  candidate ->
                      candidate.code() == AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA)
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
          List.of("A1:A2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.CellValueRule(
                      ExcelComparisonOperator.BETWEEN,
                      "1",
                      Optional.of("#REF!"),
                      false,
                      Optional.of(
                          new DifferentialStylePatch(
                              Optional.of("0.00"),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))))));

      AnalysisFindingReport finding =
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .filter(
                  candidate ->
                      candidate.code() == AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA)
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
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.FormulaRule(
                      "A1>0",
                      false,
                      Optional.of(
                          new DifferentialStylePatch(
                              Optional.of("0.00"),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))))));

      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .noneMatch(
                  finding ->
                      finding.code() == AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA));
    }
  }

  @Test
  void healthFindingsDoNotFlagHealthyCellValueRulesWithTwoValidFormulas() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.CellValueRule(
                      ExcelComparisonOperator.BETWEEN,
                      "1",
                      Optional.of("10"),
                      false,
                      Optional.of(
                          new DifferentialStylePatch(
                              Optional.of("0.00"),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))))));

      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .noneMatch(
                  finding ->
                      finding.code() == AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA));
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
              .map(AnalysisFindingReport::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE,
                      AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION)));
    }
  }

  @Test
  void healthFindingsFlagBlocksThatMixValidAndInvalidRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedConditionalFormattingValues(sheet);
      controller.setConditionalFormatting(
          sheet, List.of(new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:A3")), primaryBlock());
      ctBlock(sheet, "A1:A3").setSqref(List.of("A1:A3", "not-a-range"));

      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .contains(AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE));
    }
  }

  @Test
  void healthFindingsTraverseTop10RulesWithoutReportingBrokenFormulaNoise() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setConditionalFormatting(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList(),
          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
              List.of(
                  new ConditionalFormattingRuleSpec.Top10Rule(
                      false,
                      5,
                      false,
                      false,
                      Optional.of(
                          new DifferentialStylePatch(
                              Optional.of("0.00"),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))))));

      assertTrue(
          controller.conditionalFormattingHealthFindings("Ops", sheet).stream()
              .noneMatch(
                  finding ->
                      finding.code() == AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA));
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

      XSSFConditionalFormattingRule top10Family =
          formatting.createConditionalFormattingRule("E1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("E1:E3")}, top10Family);
      var top10FamilyCtRule = ctBlock(sheet, "E1:E3").getCfRuleArray(0);
      top10FamilyCtRule.setType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType.TOP_10);
      ConditionalFormattingRuleReport.Top10Rule top10 =
          assertInstanceOf(
              ConditionalFormattingRuleReport.Top10Rule.class,
              ExcelConditionalFormattingController.toSnapshot(
                  sheet, formatting.getConditionalFormattingAt(4).getRule(0), top10FamilyCtRule));
      assertEquals(10, top10.rank());
      assertFalse(top10.percent());
      assertFalse(top10.bottom());

      XSSFConditionalFormattingRule explicitTop10 =
          formatting.createConditionalFormattingRule("F1>0");
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("F1:F3")}, explicitTop10);
      var explicitTop10CtRule = ctBlock(sheet, "F1:F3").getCfRuleArray(0);
      explicitTop10CtRule.setType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType.TOP_10);
      explicitTop10CtRule.setRank(7L);
      explicitTop10CtRule.setPercent(true);
      explicitTop10CtRule.setBottom(true);
      explicitTop10CtRule.setStopIfTrue(true);
      ConditionalFormattingRuleReport.Top10Rule explicitTop10Snapshot =
          assertInstanceOf(
              ConditionalFormattingRuleReport.Top10Rule.class,
              ExcelConditionalFormattingController.toSnapshot(
                  sheet, formatting.getConditionalFormattingAt(5).getRule(0), explicitTop10CtRule));
      assertEquals(7, explicitTop10Snapshot.rank());
      assertTrue(explicitTop10Snapshot.percent());
      assertTrue(explicitTop10Snapshot.bottom());
      assertTrue(explicitTop10Snapshot.stopIfTrue());

      XSSFConditionalFormattingRule iconSetCatchRule =
          formatting.createConditionalFormattingRule(IconMultiStateFormatting.IconSet.GYR_3_ARROW);
      formatting.addConditionalFormatting(
          new CellRangeAddress[] {CellRangeAddress.valueOf("G1:G3")}, iconSetCatchRule);
      var iconSetCatchCtRule = ctBlock(sheet, "G1:G3").getCfRuleArray(0);
      iconSetCatchCtRule
          .getIconSet()
          .getCfvoArray(0)
          .setType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfvoType.FORMULA);
      iconSetCatchCtRule.getIconSet().getCfvoArray(0).setVal(" ");
      assertUnsupportedRule(
          ExcelConditionalFormattingController.toSnapshot(
              sheet, formatting.getConditionalFormattingAt(6).getRule(0), iconSetCatchCtRule),
          "ICON_SET",
          "unsupported threshold or icon-set payload");
    }
  }

  @Test
  void excelSheetDelegatesConditionalFormattingHealthAnalysis() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.cells().setCell("A1", new CellValue.NumberValue(1));
      sheet
          .metadata()
          .setConditionalFormatting(
              List.of("A1:A2").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                  .toList(),
              new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                  List.of(
                      new ConditionalFormattingRuleSpec.FormulaRule(
                          "#REF!",
                          false,
                          Optional.of(
                              new DifferentialStylePatch(
                                  Optional.of("0.00"),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty()))))));

      assertEquals(
          List.of(AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA),
          sheet.diagnostics().conditionalFormattingHealthFindings().stream()
              .map(AnalysisFindingReport::code)
              .distinct()
              .toList());
    }
  }

  private static dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec primaryBlock() {
    return new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
        List.of(
            new ConditionalFormattingRuleSpec.FormulaRule(
                "A1>10",
                true,
                Optional.of(
                    new DifferentialStylePatch(
                        Optional.of("0.00"),
                        Optional.of(true),
                        Optional.of(false),
                        Optional.ofNullable(
                            new dev.erst.gridgrind.core.model.style.FontHeightSpec.Points(
                                BigDecimal.valueOf(12))),
                        Optional.of(color("#102030")),
                        Optional.of(true),
                        Optional.of(true),
                        Optional.of(color("#E0F0AA")),
                        Optional.ofNullable(
                            new DifferentialBorderPatch(
                                Optional.of(borderSide(ExcelBorderStyle.THIN, "#405060")),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty()))))),
            new ConditionalFormattingRuleSpec.CellValueRule(
                ExcelComparisonOperator.BETWEEN,
                "1",
                Optional.of("9"),
                false,
                Optional.of(
                    new DifferentialStylePatch(
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of(true),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of(color("#AAEECC")),
                        Optional.empty())))));
  }

  private static dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec secondaryBlock(
      String range) {
    return new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
        List.of(
            new ConditionalFormattingRuleSpec.FormulaRule(
                range.startsWith("B") ? "B1>5" : "C1=\"Ready\"",
                false,
                Optional.of(
                    new DifferentialStylePatch(
                        Optional.empty(),
                        Optional.of(false),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of(color("#223344")),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))));
  }

  private static DifferentialBorderPatch expandedBorder(BorderSidePatch side) {
    return new DifferentialBorderPatch(
        Optional.empty(),
        Optional.of(side),
        Optional.of(side),
        Optional.of(side),
        Optional.of(side));
  }

  private static BorderSidePatch borderSide(ExcelBorderStyle style, String rgb) {
    return new BorderSidePatch(Optional.of(style), Optional.of(color(rgb)));
  }

  private static ColorSpec color(String rgb) {
    return ColorSpec.rgb(new RgbHex(rgb));
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
      ConditionalFormattingRuleReport snapshot, String kind, String detailFragment) {
    ConditionalFormattingRuleReport.UnsupportedRule unsupported =
        assertInstanceOf(ConditionalFormattingRuleReport.UnsupportedRule.class, snapshot);
    assertEquals(kind, unsupported.kind());
    assertTrue(unsupported.detail().contains(detailFragment), unsupported.detail());
  }
}
