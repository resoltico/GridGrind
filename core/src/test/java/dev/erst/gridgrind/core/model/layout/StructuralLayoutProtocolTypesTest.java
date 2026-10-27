package dev.erst.gridgrind.core.model.layout;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintAreaReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingReport;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorReport;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import org.junit.jupiter.api.Test;

/** Direct tests for the structural-layout protocol input and report families. */
class StructuralLayoutProtocolTypesTest {
  @Test
  void rowAndColumnSpanInputsValidateBounds() {
    assertEquals(
        new RowBandSelector.Span(new SheetName("Budget"), 0, 2),
        new RowBandSelector.Span(new SheetName("Budget"), 0, 2));
    assertEquals(
        new ColumnBandSelector.Span(new SheetName("Budget"), 1, 3),
        new ColumnBandSelector.Span(new SheetName("Budget"), 1, 3));
    assertEquals(
        new RowBandSelector.Insertion(new SheetName("Budget"), 2, 1),
        new RowBandSelector.Insertion(new SheetName("Budget"), 2, 1));
    assertEquals(
        new ColumnBandSelector.Insertion(new SheetName("Budget"), 3, 2),
        new ColumnBandSelector.Insertion(new SheetName("Budget"), 3, 2));

    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Span(new SheetName("Budget"), -1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Span(new SheetName("Budget"), 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Span(new SheetName("Budget"), 2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RowBandSelector.Span(
                new SheetName("Budget"), Limit.MAX_ROW_INDEX + 1, Limit.MAX_ROW_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Span(new SheetName("Budget"), 0, Limit.MAX_ROW_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Span(new SheetName("Budget"), -1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Span(new SheetName("Budget"), 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Span(new SheetName("Budget"), 2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ColumnBandSelector.Span(
                new SheetName("Budget"), Limit.MAX_COLUMN_INDEX + 1, Limit.MAX_COLUMN_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Span(new SheetName("Budget"), 0, Limit.MAX_COLUMN_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Insertion(new SheetName("Budget"), 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Insertion(new SheetName("Budget"), 0, 0));
  }

  @Test
  void paneInputVariantsValidateTheirContracts() {
    assertEquals(new PaneSpec.None(), new PaneSpec.None());
    assertEquals(new PaneSpec.Frozen(1, 2, 1, 2), new PaneSpec.Frozen(1, 2, 1, 2));
    assertEquals(
        new PaneSpec.Split(0, 1800, 0, 3, ExcelPaneRegion.LOWER_LEFT),
        new PaneSpec.Split(0, 1800, 0, 3, ExcelPaneRegion.LOWER_LEFT));

    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(-1, 2, 1, 2));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneSpec.Split(0, 0, 0, 0, ExcelPaneRegion.UPPER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneSpec.Split(0, 1800, 1, 3, ExcelPaneRegion.LOWER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneSpec.Split(1200, 0, 2, 1, ExcelPaneRegion.UPPER_RIGHT));
    assertThrows(NullPointerException.class, () -> new PaneSpec.Split(1200, 1800, 2, 3, null));
  }

  @Test
  void printAreaAndScalingInputsCoverDefaultsAndBounds() {
    assertEquals(new PrintAreaSpec.None(), new PrintAreaSpec.None());
    assertEquals(
        new PrintAreaSpec.Range(new RangeRef("A1:C10")),
        new PrintAreaSpec.Range(new RangeRef("A1:C10")));
    assertThrows(IllegalArgumentException.class, () -> new PrintAreaSpec.Range(new RangeRef(" ")));

    assertEquals(new PrintScalingSpec.Automatic(), new PrintScalingSpec.Automatic());
    assertEquals(new PrintScalingSpec.Fit(1, 0), new PrintScalingSpec.Fit(1, 0));
    assertThrows(IllegalArgumentException.class, () -> new PrintScalingSpec.Fit(-1, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new PrintScalingSpec.Fit(Short.MAX_VALUE + 1, 1));
  }

  @Test
  void printTitleInputsCoverNoneBandsAndBounds() {
    assertEquals(new PrintTitleRowsSpec.None(), new PrintTitleRowsSpec.None());
    assertEquals(new PrintTitleRowsSpec.Band(0, 1), new PrintTitleRowsSpec.Band(0, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(-1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintTitleRowsSpec.Band(0, Limit.MAX_ROW_INDEX + 1));

    assertEquals(new PrintTitleColumnsSpec.None(), new PrintTitleColumnsSpec.None());
    assertEquals(new PrintTitleColumnsSpec.Band(0, 1), new PrintTitleColumnsSpec.Band(0, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(-1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintTitleColumnsSpec.Band(0, Limit.MAX_COLUMN_INDEX + 1));
  }

  @Test
  void printLayoutInputNormalizesNullsToSupportedDefaults() {
    PrintLayoutSpec input = PrintLayoutSpec.defaults();

    assertEquals(new PrintAreaSpec.None(), input.printArea());
    assertEquals(ExcelPrintOrientation.PORTRAIT, input.orientation());
    assertEquals(new PrintScalingSpec.Automatic(), input.scaling());
    assertEquals(new PrintTitleRowsSpec.None(), input.repeatingRows());
    assertEquals(new PrintTitleColumnsSpec.None(), input.repeatingColumns());
    assertEquals(HeaderFooterTextSpec.blank(), input.header());
    assertEquals(HeaderFooterTextSpec.blank(), input.footer());
  }

  @Test
  void paneAndPrintReportVariantsValidateTheirContracts() {
    assertEquals(new PaneReport.None(), new PaneReport.None());
    assertEquals(
        new PaneReport.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
        new PaneReport.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT));
    assertThrows(NullPointerException.class, () -> new PaneReport.Split(1, 1, 0, 0, null));

    assertEquals(new PrintAreaReport.None(), new PrintAreaReport.None());
    assertEquals(
        new PrintAreaReport.Range(new RangeLabel("A1:B20")),
        new PrintAreaReport.Range(new RangeLabel("A1:B20")));
    assertDoesNotThrow(() -> new PrintAreaReport.Range(new RangeLabel(" ")));
    assertEquals(new PrintScalingReport.Automatic(), new PrintScalingReport.Automatic());
    assertEquals(new PrintTitleRowsReport.None(), new PrintTitleRowsReport.None());
    assertEquals(new PrintTitleColumnsReport.None(), new PrintTitleColumnsReport.None());
    assertEquals(
        new HeaderFooterTextReport("Left", "Center", "Right"),
        new HeaderFooterTextReport("Left", "Center", "Right"));
    assertThrows(NullPointerException.class, () -> new HeaderFooterTextReport(null, "", ""));

    assertEquals(
        new PrintLayoutReport(
            new SheetLabel("Budget"),
            new PrintAreaReport.None(),
            ExcelPrintOrientation.PORTRAIT,
            new PrintScalingReport.Automatic(),
            new PrintTitleRowsReport.None(),
            new PrintTitleColumnsReport.None(),
            new HeaderFooterTextReport("", "", ""),
            new HeaderFooterTextReport("", "", "")),
        new PrintLayoutReport(
            new SheetLabel("Budget"),
            new PrintAreaReport.None(),
            ExcelPrintOrientation.PORTRAIT,
            new PrintScalingReport.Automatic(),
            new PrintTitleRowsReport.None(),
            new PrintTitleColumnsReport.None(),
            new HeaderFooterTextReport("", "", ""),
            new HeaderFooterTextReport("", "", "")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintLayoutReport(
                new SheetLabel(" "),
                new PrintAreaReport.None(),
                ExcelPrintOrientation.PORTRAIT,
                new PrintScalingReport.Automatic(),
                new PrintTitleRowsReport.None(),
                new PrintTitleColumnsReport.None(),
                new HeaderFooterTextReport("", "", ""),
                new HeaderFooterTextReport("", "", "")));
  }

  @Test
  void sheetLayoutReportRejectsMissingPaneAndOutOfBoundsZoom() {
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetLayoutReport(
                new SheetLabel("Budget"),
                null,
                100,
                SheetPresentationReport.defaults(),
                java.util.List.of(),
                java.util.List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetLayoutReport(
                new SheetLabel("Budget"),
                new PaneReport.None(),
                9,
                SheetPresentationReport.defaults(),
                java.util.List.of(),
                java.util.List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetLayoutReport(
                new SheetLabel("Budget"),
                new PaneReport.None(),
                401,
                SheetPresentationReport.defaults(),
                java.util.List.of(),
                java.util.List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetLayoutReport(
                new SheetLabel("Budget"),
                new PaneReport.None(),
                100,
                null,
                java.util.List.of(),
                java.util.List.of()));
  }

  @Test
  void sheetPresentationInputsAndReportsNormalizeAndValidate() {
    SheetPresentationSpec defaultPresentation = SheetPresentationSpec.defaults();
    SheetPresentationSpec explicitPresentation =
        new SheetPresentationSpec(
            new SheetDisplaySpec(false, false, false, true, true),
            java.util.Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
            new SheetOutlineSummarySpec(false, false),
            new SheetDefaultsSpec(11, 18.5d),
            java.util.List.of(
                new IgnoredErrorSpec(
                    new RangeRef("A1:B2"),
                    java.util.List.of(
                        dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType.FORMULA))));

    assertEquals(SheetDisplaySpec.defaults(), defaultPresentation.display());
    assertEquals(java.util.Optional.empty(), defaultPresentation.tabColor());
    assertEquals(SheetOutlineSummarySpec.defaults(), defaultPresentation.outlineSummary());
    assertEquals(SheetDefaultsSpec.defaults(), defaultPresentation.sheetDefaults());
    assertEquals(java.util.List.of(), defaultPresentation.ignoredErrors());
    assertEquals(
        new SheetDisplaySpec(false, false, false, true, true), explicitPresentation.display());
    assertEquals(
        java.util.Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
        explicitPresentation.tabColor());
    assertEquals(new SheetOutlineSummarySpec(false, false), explicitPresentation.outlineSummary());
    assertEquals(new SheetDefaultsSpec(11, 18.5d), explicitPresentation.sheetDefaults());
    assertEquals(
        java.util.List.of(
            new IgnoredErrorSpec(
                new RangeRef("A1:B2"), java.util.List.of(ExcelIgnoredErrorType.FORMULA))),
        explicitPresentation.ignoredErrors());
    assertEquals(
        new SheetPresentationReport(
            SheetDisplayReport.defaults(),
            java.util.Optional.empty(),
            SheetOutlineSummaryReport.defaults(),
            SheetDefaultsReport.defaults(),
            java.util.List.of()),
        SheetPresentationReport.defaults());
    assertEquals(
        new SheetPresentationReport(
            new SheetDisplayReport(false, false, false, true, true),
            java.util.Optional.of(CellColorReport.rgb(new RgbHex("#112233"))),
            new SheetOutlineSummaryReport(false, false),
            new SheetDefaultsReport(11, 18.5d),
            java.util.List.of(
                new IgnoredErrorReport(
                    new RangeLabel("A1:B2"), java.util.List.of(ExcelIgnoredErrorType.FORMULA)))),
        new SheetPresentationReport(
            new SheetDisplayReport(false, false, false, true, true),
            java.util.Optional.of(CellColorReport.rgb(new RgbHex("#112233"))),
            new SheetOutlineSummaryReport(false, false),
            new SheetDefaultsReport(11, 18.5d),
            java.util.List.of(
                new IgnoredErrorReport(
                    new RangeLabel("A1:B2"), java.util.List.of(ExcelIgnoredErrorType.FORMULA)))));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsSpec(0, 15.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new SheetDefaultsSpec(Limit.MAX_DEFAULT_COLUMN_WIDTH + 1, 15.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new SheetDefaultsSpec(8, Math.nextUp(Limit.MAX_ROW_HEIGHT_POINTS)));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsReport(8, 0.0d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetPresentationSpec(
                SheetDisplaySpec.defaults(),
                java.util.Optional.empty(),
                SheetOutlineSummarySpec.defaults(),
                SheetDefaultsSpec.defaults(),
                java.util.List.of(
                    new IgnoredErrorSpec(
                        new RangeRef("A1"), java.util.List.of(ExcelIgnoredErrorType.FORMULA)),
                    new IgnoredErrorSpec(
                        new RangeRef("A1"),
                        java.util.List.of(ExcelIgnoredErrorType.EVALUATION_ERROR)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new IgnoredErrorSpec(
                new RangeRef("A1"),
                java.util.List.of(ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.FORMULA)));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                java.util.Optional.empty(),
                SheetOutlineSummaryReport.defaults(),
                SheetDefaultsReport.defaults(),
                null));
    assertEquals(
        "tabColor must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new SheetPresentationReport(
                        SheetDisplayReport.defaults(),
                        null,
                        SheetOutlineSummaryReport.defaults(),
                        SheetDefaultsReport.defaults(),
                        java.util.List.of(
                            new IgnoredErrorReport(
                                new RangeLabel("A1"),
                                java.util.List.of(ExcelIgnoredErrorType.FORMULA)),
                            new IgnoredErrorReport(
                                new RangeLabel("A1"),
                                java.util.List.of(ExcelIgnoredErrorType.EVALUATION_ERROR)))))
            .getMessage());
  }

  @Test
  void sheetDisplayDefaultsAndIgnoredErrorDtosCoverConstructorBranches() {
    assertEquals(SheetDisplaySpec.defaults(), new SheetDisplaySpec(true, true, true, false, false));
    assertEquals(
        new SheetDisplaySpec(false, false, true, true, false),
        new SheetDisplaySpec(false, false, true, true, false));
    assertEquals(SheetOutlineSummarySpec.defaults(), new SheetOutlineSummarySpec(true, true));
    assertEquals(
        new SheetOutlineSummarySpec(true, false), new SheetOutlineSummarySpec(true, false));
    assertEquals(SheetDefaultsSpec.defaults(), new SheetDefaultsSpec(8, 15.0d));
    assertEquals(new SheetDefaultsSpec(8, 18.5d), new SheetDefaultsSpec(8, 18.5d));
    assertEquals(SheetDefaultsReport.defaults(), new SheetDefaultsReport(8, 15.0d));

    java.util.List<ExcelIgnoredErrorType> errorTypesWithNull =
        new java.util.ArrayList<>(java.util.Arrays.asList(ExcelIgnoredErrorType.FORMULA, null));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsSpec(8, 0.0d));
    assertThrows(
        IllegalArgumentException.class, () -> new SheetDefaultsSpec(8, Double.POSITIVE_INFINITY));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsReport(0, 15.0d));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsReport(8, Double.NaN));
    assertThrows(
        NullPointerException.class,
        () -> new IgnoredErrorSpec(null, java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new IgnoredErrorSpec(
                new RangeRef(" "), java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertThrows(NullPointerException.class, () -> new IgnoredErrorSpec(new RangeRef("A1"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new IgnoredErrorSpec(new RangeRef("A1"), java.util.List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new IgnoredErrorSpec(new RangeRef("A1"), errorTypesWithNull));
    assertThrows(
        NullPointerException.class,
        () -> new IgnoredErrorReport(null, java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertDoesNotThrow(
        () ->
            new IgnoredErrorReport(
                new RangeLabel(" "), java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertThrows(
        NullPointerException.class, () -> new IgnoredErrorReport(new RangeLabel("A1"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new IgnoredErrorReport(new RangeLabel("A1"), java.util.List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new IgnoredErrorReport(new RangeLabel("A1"), errorTypesWithNull));
  }
}
