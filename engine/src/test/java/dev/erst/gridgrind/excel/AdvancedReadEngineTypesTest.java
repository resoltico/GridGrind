package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientStopReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.plan.ResolvedInspectionStep;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Tests for richer factual engine snapshot types added for advanced XSSF parity reads. */
class AdvancedReadEngineTypesTest {
  @Test
  void colorGradientAndFillSnapshotsValidateRichReadShapes() {
    assertEquals(
        CellColorReport.rgb(new RgbHex("#ABCDEF")), CellColorReport.rgb(new RgbHex("#abcdef")));
    assertEquals(4, CellColorReport.theme(4, 0.45d).theme());
    assertThrows(NullPointerException.class, () -> CellColorReport.rgb(null));
    assertThrows(IllegalArgumentException.class, () -> CellColorReport.rgb(new RgbHex(" ")));
    assertThrows(IllegalArgumentException.class, () -> CellColorReport.rgb(new RgbHex("#12345G")));
    assertThrows(IllegalArgumentException.class, () -> CellColorReport.theme(-1));
    assertThrows(IllegalArgumentException.class, () -> CellColorReport.indexed(-1));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellColorReport.rgb(new RgbHex("#112233"), Double.NaN));

    CellGradientStopReport firstStop = new CellGradientStopReport(0.0d, rgb("#112233"));
    CellGradientStopReport secondStop =
        new CellGradientStopReport(1.0d, CellColorReport.theme(4, 0.45d));
    CellGradientFillReport gradient =
        CellGradientFillReport.path(0.1d, 0.2d, 0.3d, 0.4d, List.of(firstStop, secondStop));

    assertEquals(2, gradient.stops().size());
    assertEquals(0.2d, gradientRight(gradient));
    assertEquals(
        CellFillReport.patternForeground(ExcelFillPattern.SOLID, rgb("#112233")),
        CellFillReport.patternForeground(ExcelFillPattern.SOLID, rgb("#112233")));
    assertEquals(
        CellFillReport.pattern(ExcelFillPattern.NONE),
        CellFillReport.pattern(ExcelFillPattern.NONE));
    CellFillReport.PatternBackground backgroundOnlyFill =
        assertInstanceOf(
            CellFillReport.PatternBackground.class,
            CellFillReport.patternBackground(ExcelFillPattern.BRICKS, rgb("#445566")));
    assertEquals(ExcelFillPattern.BRICKS, backgroundOnlyFill.pattern());
    assertEquals(rgb("#445566"), backgroundOnlyFill.backgroundColor());
    CellGradientFillReport.Linear implicitLinearGradient =
        assertInstanceOf(
            CellGradientFillReport.Linear.class,
            CellGradientFillReport.linear(null, List.of(firstStop, secondStop)));
    assertNull(implicitLinearGradient.degree());
    assertThrows(
        IllegalArgumentException.class, () -> new CellGradientStopReport(1.5d, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class, () -> new CellGradientStopReport(-0.1d, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopReport(Double.NaN, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellGradientFillReport.path(
                Double.NaN, null, null, null, List.of(firstStop, secondStop)));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellGradientFillReport.linear(Double.POSITIVE_INFINITY, List.of(firstStop)));
    assertThrows(
        IllegalArgumentException.class, () -> CellGradientFillReport.linear(45.0d, List.of()));
    assertThrows(
        NullPointerException.class,
        () -> CellGradientFillReport.linear(45.0d, Arrays.asList(firstStop, null)));

    assertEquals(gradient, CellFillReport.gradient(gradient).gradient());
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillReport.patternForeground(ExcelFillPattern.NONE, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillReport.patternColors(ExcelFillPattern.SOLID, rgb("#112233"), rgb("#445566")));
    assertThrows(NullPointerException.class, () -> CellFillReport.gradient(null));
    assertThrows(
        NullPointerException.class,
        () -> CellFillReport.patternBackground(ExcelFillPattern.BRICKS, null));
  }

  @Test
  void colorSnapshotSupportAndRgbHelpersPreserveWorkbookColorSemantics() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFColor rgbColor =
          new XSSFColor(
              new byte[] {0x11, 0x22, 0x33}, workbook.getStylesSource().getIndexedColors());
      XSSFColor indexedColor = new XSSFColor();
      indexedColor.setIndexed(IndexedColors.DARK_RED.getIndex());
      XSSFColor tintedTheme = new XSSFColor();
      tintedTheme.setTheme(3);
      tintedTheme.setTint(0.25d);
      XSSFColor malformedRgb =
          new XSSFColor(new byte[] {0x11, 0x22}, workbook.getStylesSource().getIndexedColors());
      CTColor themed = CTColor.Factory.newInstance();
      themed.setTheme(2L);
      CTColor indexed = CTColor.Factory.newInstance();
      indexed.setIndexed(7L);
      CTColor argb = CTColor.Factory.newInstance();
      argb.setRgb(new byte[] {(byte) 0xFF, 0x44, 0x55, 0x66});

      assertEquals(Optional.empty(), ExcelColorSnapshotSupport.snapshot((XSSFColor) null));
      assertThrows(
          IllegalStateException.class, () -> ExcelColorSnapshotSupport.snapshot(new XSSFColor()));
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#112233"))),
          ExcelColorSnapshotSupport.snapshot(rgbColor));
      assertEquals(
          Optional.of(CellColorReport.indexed((int) IndexedColors.DARK_RED.getIndex())),
          ExcelColorSnapshotSupport.snapshot(indexedColor));
      assertEquals(
          Optional.of(CellColorReport.theme(3, 0.25d)),
          ExcelColorSnapshotSupport.snapshot(tintedTheme));
      assertEquals(Optional.of("#112233"), ExcelRgbColorSupport.toRgbHex(rgbColor));
      assertEquals(Optional.empty(), ExcelRgbColorSupport.toRgbHex(null));
      assertEquals(Optional.empty(), ExcelRgbColorSupport.toRgbHex(tintedTheme));
      assertEquals(Optional.empty(), ExcelRgbColorSupport.toRgbHex(malformedRgb));
      assertEquals(Optional.empty(), ExcelColorSnapshotSupport.snapshot(workbook, null));
      assertEquals(
          Optional.of(CellColorReport.theme(2)),
          ExcelColorSnapshotSupport.snapshot(workbook, themed));
      assertEquals(
          Optional.of(CellColorReport.indexed(7)),
          ExcelColorSnapshotSupport.snapshot(workbook, indexed));
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#445566"))),
          ExcelColorSnapshotSupport.snapshot(workbook, argb));
      assertNull(ExcelColorSupport.copyOf(null));
      assertEquals(
          ColorSpec.theme(3, 0.25d), ExcelColorSupport.copyOf(CellColorReport.theme(3, 0.25d)));
      assertEquals(ColorSpec.indexed(7), ExcelColorSupport.copyOf(CellColorReport.indexed(7)));
      assertThrows(
          IllegalStateException.class,
          () -> ExcelColorSnapshotSupport.snapshot(workbook, CTColor.Factory.newInstance()));
    }
  }

  @Test
  void printAndProtectionReportsValidateRichReadShapes() {
    PrintSetupReport setup =
        new PrintSetupReport(
            new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
            false,
            true,
            false,
            9,
            false,
            true,
            2,
            true,
            3,
            List.of(10, 20),
            List.of(2, 4));

    assertEquals(List.of(10, 20), setup.rowBreaks());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsReport(-0.1d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsReport(Double.POSITIVE_INFINITY, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                -1,
                false,
                true,
                2,
                true,
                3,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                -1,
                true,
                3,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                -1,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                List.of(-1),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                List.of(10),
                List.of(-1)));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                Arrays.asList(10, null),
                List.of(2)));

    WorkbookProtectionReport protection =
        new WorkbookProtectionReport(true, false, true, true, false);
    ResolvedInspectionStep read =
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("workbook-protection"),
            new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
            new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetWorkbookProtection());
    dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookProtectionResult result =
        new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookProtectionResult(
            new StepId("workbook-protection"), protection);

    assertTrue(protection.structureLocked());
    assertEquals("workbook-protection", read.stepId().value());
    assertTrue(result.protection().workbookPasswordHashPresent());
    assertThrows(
        NullPointerException.class,
        () ->
            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                null,
                new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                    .GetWorkbookProtection()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                new StepId(" "),
                new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                    .GetWorkbookProtection()));
    assertThrows(
        NullPointerException.class,
        () ->
            new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookProtectionResult(
                new StepId("workbook-protection"), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookProtectionResult(
                new StepId(" "), new WorkbookProtectionReport(true, false, true, true, false)));
  }

  @Test
  void autofilterAndTableSnapshotsValidateAdvancedCriteriaAndSortShapes() {
    AutofilterFilterCriterionReport.Values values =
        new AutofilterFilterCriterionReport.Values(List.of("Queued", "Blocked"), true);
    AutofilterFilterCriterionReport.Custom custom =
        new AutofilterFilterCriterionReport.Custom(
            true,
            List.of(new AutofilterFilterCriterionReport.CustomConditionReport("equal", "Ada")));
    AutofilterFilterCriterionReport.Dynamic dynamic =
        new AutofilterFilterCriterionReport.Dynamic("TODAY", 1.0d, 2.0d);
    AutofilterFilterCriterionReport.Top10 top10 =
        new AutofilterFilterCriterionReport.Top10(true, false, 10.0d, 8.0d);
    AutofilterFilterCriterionReport.Color color =
        new AutofilterFilterCriterionReport.Color(false, CellColorReport.theme(4, 0.45d));
    AutofilterFilterCriterionReport.Icon icon =
        new AutofilterFilterCriterionReport.Icon("3TrafficLights1", 2);
    AutofilterSortConditionReport sortCondition =
        new AutofilterSortConditionReport.CellColor(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A5"), true, rgb("#AABBCC"));
    AutofilterSortStateReport sortState =
        new AutofilterSortStateReport(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F5"),
            true,
            false,
            java.util.Optional.empty(),
            List.of(sortCondition));
    AutofilterEntryReport.SheetOwned sheetOwned =
        new AutofilterEntryReport.SheetOwned(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F5"),
            List.of(
                new AutofilterFilterColumnReport(0L, false, values),
                new AutofilterFilterColumnReport(1L, true, custom),
                new AutofilterFilterColumnReport(2L, true, dynamic),
                new AutofilterFilterColumnReport(3L, true, top10),
                new AutofilterFilterColumnReport(4L, true, color),
                new AutofilterFilterColumnReport(5L, true, icon)),
            java.util.Optional.of(sortState));
    AutofilterEntryReport.TableOwned tableOwned =
        new AutofilterEntryReport.TableOwned(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("H1:I5"),
            "QueueTable",
            List.of(),
            java.util.Optional.of(sortState));
    AutofilterEntryReport.TableOwned tableOwnedWithColumn =
        new AutofilterEntryReport.TableOwned(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("N1:O5"),
            "QueueMirror",
            List.of(new AutofilterFilterColumnReport(9L, true, values)),
            java.util.Optional.of(sortState));
    AutofilterEntryReport.TableOwned defaultTableOwned =
        new AutofilterEntryReport.TableOwned(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("L1:M4"), "AuditTable");
    AutofilterFilterCriterionReport.Values emptyValues =
        new AutofilterFilterCriterionReport.Values(List.of(), false);

    assertEquals(java.util.Optional.of(sortState), sheetOwned.sortState());
    assertEquals("QueueTable", tableOwned.tableName());
    assertEquals(1, tableOwnedWithColumn.filterColumns().size());
    assertEquals(List.of(), emptyValues.values());
    assertEquals(List.of(), defaultTableOwned.filterColumns());
    assertTrue(defaultTableOwned.sortState().isEmpty());
    assertInstanceOf(AutofilterSortConditionReport.CellColor.class, sortCondition);
    assertEquals(java.util.Optional.empty(), sortState.sortMethod());

    assertThrows(
        NullPointerException.class,
        () -> new AutofilterFilterCriterionReport.Values(Arrays.asList("Queued", null), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Custom(true, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.CustomConditionReport(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.CustomConditionReport("equal", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic(" ", 1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic("TODAY", Double.POSITIVE_INFINITY, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic("TODAY", 1.0d, Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Top10(true, false, -1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Top10(true, false, Double.NaN, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterFilterCriterionReport.Top10(
                true, false, 10.0d, Double.NEGATIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterCriterionReport.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterColumnReport(-1L, true, values));
    assertInstanceOf(
        AutofilterSortConditionReport.CellColor.class,
        new AutofilterSortConditionReport.CellColor(
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A5"), false, rgb("#AABBCC")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterSortConditionReport.Icon(
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A5"), true, -1));
    assertDoesNotThrow(
        () ->
            new AutofilterSortConditionReport.Value(
                new dev.erst.gridgrind.core.vocabulary.RangeLabel(" "), false));
    assertDoesNotThrow(
        () ->
            new AutofilterSortStateReport(
                new dev.erst.gridgrind.core.vocabulary.RangeLabel(" "),
                true,
                false,
                java.util.Optional.empty(),
                List.of(sortCondition)));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterSortStateReport(
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F5"),
                true,
                false,
                java.util.Optional.empty(),
                Arrays.asList(sortCondition, null)));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterEntryReport.TableOwned(
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:F5"),
                "QueueTable",
                Arrays.asList(sheetOwned.filterColumns().getFirst(), null),
                java.util.Optional.of(sortState)));
  }

  @Test
  void tableAndConditionalFormattingSnapshotsValidateAdvancedMetadata() {
    TableEntryReport normalized =
        new TableEntryReport(
            new dev.erst.gridgrind.core.vocabulary.TableLabel("BudgetTable"),
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B5"),
            new TableEntryReport.Structure(
                1,
                1,
                List.of("Item", "Amount"),
                List.of(
                    new TableColumnReport(
                        1L,
                        "Item",
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()),
                    new TableColumnReport(
                        2L,
                        "Amount",
                        Optional.ofNullable("UniqueAmount"),
                        Optional.ofNullable("Total"),
                        Optional.ofNullable("sum"),
                        Optional.ofNullable("[@Amount]*2")))),
            new TableStyleReport.Named("TableStyleMedium2", false, false, true, false),
            new TableEntryReport.Behavior(true, true, true, false),
            new TableEntryReport.Presentation(
                Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()));
    ConditionalFormattingRuleReport.Top10Rule top10 =
        new ConditionalFormattingRuleReport.Top10Rule(
            1,
            false,
            10,
            true,
            false,
            Optional.ofNullable(
                new DifferentialStyleReport(
                    "0.00",
                    true,
                    null,
                    null,
                    Optional.ofNullable(
                        dev.erst.gridgrind.excel.DifferentialStyleFixtures.color(
                            ColorSpec.rgb(new RgbHex("#AABBCC")))),
                    null,
                    null,
                    Optional.empty(),
                    Optional.empty(),
                    List.of())));

    assertEquals(Optional.empty(), normalized.presentation().comment());
    assertEquals(Optional.empty(), normalized.presentation().headerRowCellStyle());
    assertEquals(
        "UniqueAmount", normalized.structure().columns().get(1).uniqueName().orElseThrow());
    assertEquals(10, top10.rank());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableColumnReport(
                -1L,
                "Item",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.Top10Rule(
                1, false, -1, false, false, Optional.empty()));
  }

  private static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(new RgbHex(rgb));
  }
}
