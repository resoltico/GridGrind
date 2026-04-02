package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for sheet-copy planning and validation helpers. */
class ExcelSheetCopyControllerTest {
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
  void localNamedRangeHelpersRespectScopeBoundaries() {
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
                "WorkbookFormula", new ExcelNamedRangeScope.WorkbookScope(), "SUM(Source!$A$1)"));

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
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelSheetCopyController.requireNoUncopyableLocalNamedRanges(
                    List.of(
                        new ExcelNamedRangeSnapshot.FormulaSnapshot(
                            "LocalFormula",
                            new ExcelNamedRangeScope.SheetScope("Source"),
                            "SUM(Source!$A$1:$A$2)")),
                    "Source"));
    assertEquals(
        "cannot copy sheet 'Source': sheet-scoped formula-defined named ranges are not copyable",
        failure.getMessage());
  }

  @Test
  void supportedDataValidationsRejectUnsupportedRules() {
    ExcelDataValidationSnapshot.Supported supported =
        new ExcelDataValidationSnapshot.Supported(List.of("A1"), validationDefinition());

    assertEquals(
        List.of(supported),
        ExcelSheetCopyController.supportedDataValidations(List.of(supported), "Source"));
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelSheetCopyController.supportedDataValidations(
                    List.of(
                        new ExcelDataValidationSnapshot.Unsupported(
                            List.of("B2"), "CUSTOM", "detail")),
                    "Source"));
    assertEquals(
        "cannot copy sheet 'Source': unsupported data validation 'CUSTOM' is not copyable",
        failure.getMessage());
  }

  @Test
  void supportedConditionalFormattingRejectsUncopyableRuleFamilies() {
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
    assertEquals(
        "cannot copy sheet 'Source': conditional-formatting color scales are not copyable",
        assertThrows(
                IllegalArgumentException.class,
                () ->
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
                        "Source"))
            .getMessage());
    assertEquals(
        "cannot copy sheet 'Source': conditional-formatting data bars are not copyable",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    ExcelSheetCopyController.copyableRule(
                        new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                            1,
                            false,
                            "#102030",
                            false,
                            true,
                            0,
                            100,
                            new ExcelConditionalFormattingThresholdSnapshot(
                                ExcelConditionalFormattingThresholdType.MIN, null, null),
                            new ExcelConditionalFormattingThresholdSnapshot(
                                ExcelConditionalFormattingThresholdType.MAX, null, null)),
                        "Source"))
            .getMessage());
    assertEquals(
        "cannot copy sheet 'Source': conditional-formatting icon sets are not copyable",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    ExcelSheetCopyController.copyableRule(
                        new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
                            1,
                            false,
                            ExcelConditionalFormattingIconSet.GYR_3_ARROW,
                            false,
                            false,
                            List.of(
                                new ExcelConditionalFormattingThresholdSnapshot(
                                    ExcelConditionalFormattingThresholdType.MIN, null, null),
                                new ExcelConditionalFormattingThresholdSnapshot(
                                    ExcelConditionalFormattingThresholdType.MAX, null, null))),
                        "Source"))
            .getMessage());
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

  private static ExcelDataValidationDefinition validationDefinition() {
    return new ExcelDataValidationDefinition(
        new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
        true,
        false,
        null,
        null);
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
