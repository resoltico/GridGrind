package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import org.junit.jupiter.api.Test;

/** Direct tests for the structural-layout protocol input and report families. */
class StructuralLayoutProtocolTypesTest {
  @Test
  void rowAndColumnSpanInputsValidateBounds() {
    assertEquals(
        new RowBandSelector.Span("Budget", 0, 2), new RowBandSelector.Span("Budget", 0, 2));
    assertEquals(
        new ColumnBandSelector.Span("Budget", 1, 3), new ColumnBandSelector.Span("Budget", 1, 3));
    assertEquals(
        new RowBandSelector.Insertion("Budget", 2, 1),
        new RowBandSelector.Insertion("Budget", 2, 1));
    assertEquals(
        new ColumnBandSelector.Insertion("Budget", 3, 2),
        new ColumnBandSelector.Insertion("Budget", 3, 2));

    assertThrows(IllegalArgumentException.class, () -> new RowBandSelector.Span("Budget", -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new RowBandSelector.Span("Budget", 0, -1));
    assertThrows(IllegalArgumentException.class, () -> new RowBandSelector.Span("Budget", 2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RowBandSelector.Span(
                "Budget", ExcelRowSpan.MAX_ROW_INDEX + 1, ExcelRowSpan.MAX_ROW_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Span("Budget", 0, ExcelRowSpan.MAX_ROW_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnBandSelector.Span("Budget", -1, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnBandSelector.Span("Budget", 0, -1));
    assertThrows(IllegalArgumentException.class, () -> new ColumnBandSelector.Span("Budget", 2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ColumnBandSelector.Span(
                "Budget",
                ExcelColumnSpan.MAX_COLUMN_INDEX + 1,
                ExcelColumnSpan.MAX_COLUMN_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Span("Budget", 0, ExcelColumnSpan.MAX_COLUMN_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class, () -> new RowBandSelector.Insertion("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnBandSelector.Insertion("Budget", 0, 0));
  }

  @Test
  void paneInputVariantsValidateTheirContracts() {
    assertEquals(new PaneInput.None(), new PaneInput.None());
    assertEquals(new PaneInput.Frozen(1, 2, 1, 2), new PaneInput.Frozen(1, 2, 1, 2));
    assertEquals(
        new PaneInput.Split(0, 1800, 0, 3, ExcelPaneRegion.LOWER_LEFT),
        new PaneInput.Split(0, 1800, 0, 3, ExcelPaneRegion.LOWER_LEFT));

    assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(-1, 2, 1, 2));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneInput.Split(0, 0, 0, 0, ExcelPaneRegion.UPPER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneInput.Split(0, 1800, 1, 3, ExcelPaneRegion.LOWER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneInput.Split(1200, 0, 2, 1, ExcelPaneRegion.UPPER_RIGHT));
    assertThrows(NullPointerException.class, () -> new PaneInput.Split(1200, 1800, 2, 3, null));
  }

  @Test
  void printAreaAndScalingInputsCoverDefaultsAndBounds() {
    assertEquals(new PrintAreaInput.None(), new PrintAreaInput.None());
    assertEquals(new PrintAreaInput.Range("A1:C10"), new PrintAreaInput.Range("A1:C10"));
    assertThrows(IllegalArgumentException.class, () -> new PrintAreaInput.Range(" "));

    assertEquals(new PrintScalingInput.Automatic(), new PrintScalingInput.Automatic());
    assertEquals(new PrintScalingInput.Fit(1, 0), new PrintScalingInput.Fit(1, 0));
    assertThrows(IllegalArgumentException.class, () -> new PrintScalingInput.Fit(-1, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new PrintScalingInput.Fit(Short.MAX_VALUE + 1, 1));
  }

  @Test
  void printTitleInputsCoverNoneBandsAndBounds() {
    assertEquals(new PrintTitleRowsInput.None(), new PrintTitleRowsInput.None());
    assertEquals(new PrintTitleRowsInput.Band(0, 1), new PrintTitleRowsInput.Band(0, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsInput.Band(null, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsInput.Band(0, null));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsInput.Band(-1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsInput.Band(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsInput.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintTitleRowsInput.Band(0, PrintTitleRowsInput.MAX_ROW_INDEX + 1));

    assertEquals(new PrintTitleColumnsInput.None(), new PrintTitleColumnsInput.None());
    assertEquals(new PrintTitleColumnsInput.Band(0, 1), new PrintTitleColumnsInput.Band(0, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsInput.Band(null, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsInput.Band(0, null));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsInput.Band(-1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsInput.Band(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsInput.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintTitleColumnsInput.Band(0, PrintTitleColumnsInput.MAX_COLUMN_INDEX + 1));
  }

  @Test
  void printLayoutInputNormalizesNullsToSupportedDefaults() {
    PrintLayoutInput input = new PrintLayoutInput(null, null, null, null, null, null, null);

    assertEquals(new PrintAreaInput.None(), input.printArea());
    assertEquals(ExcelPrintOrientation.PORTRAIT, input.orientation());
    assertEquals(new PrintScalingInput.Automatic(), input.scaling());
    assertEquals(new PrintTitleRowsInput.None(), input.repeatingRows());
    assertEquals(new PrintTitleColumnsInput.None(), input.repeatingColumns());
    assertEquals(HeaderFooterTextInput.blank(), input.header());
    assertEquals(HeaderFooterTextInput.blank(), input.footer());
  }

  @Test
  void paneAndPrintReportVariantsValidateTheirContracts() {
    assertEquals(new PaneReport.None(), new PaneReport.None());
    assertEquals(
        new PaneReport.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
        new PaneReport.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT));
    assertThrows(NullPointerException.class, () -> new PaneReport.Split(1, 1, 0, 0, null));

    assertEquals(new PrintAreaReport.None(), new PrintAreaReport.None());
    assertEquals(new PrintAreaReport.Range("A1:B20"), new PrintAreaReport.Range("A1:B20"));
    assertThrows(IllegalArgumentException.class, () -> new PrintAreaReport.Range(" "));
    assertEquals(new PrintScalingReport.Automatic(), new PrintScalingReport.Automatic());
    assertEquals(new PrintTitleRowsReport.None(), new PrintTitleRowsReport.None());
    assertEquals(new PrintTitleColumnsReport.None(), new PrintTitleColumnsReport.None());
    assertEquals(
        new HeaderFooterTextReport("Left", "Center", "Right"),
        new HeaderFooterTextReport("Left", "Center", "Right"));
    assertThrows(NullPointerException.class, () -> new HeaderFooterTextReport(null, "", ""));

    assertEquals(
        new PrintLayoutReport(
            "Budget",
            new PrintAreaReport.None(),
            ExcelPrintOrientation.PORTRAIT,
            new PrintScalingReport.Automatic(),
            new PrintTitleRowsReport.None(),
            new PrintTitleColumnsReport.None(),
            new HeaderFooterTextReport("", "", ""),
            new HeaderFooterTextReport("", "", "")),
        new PrintLayoutReport(
            "Budget",
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
                " ",
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
            new GridGrindResponse.SheetLayoutReport(
                "Budget",
                null,
                100,
                SheetPresentationReport.defaults(),
                java.util.List.of(),
                java.util.List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.SheetLayoutReport(
                "Budget",
                new PaneReport.None(),
                9,
                SheetPresentationReport.defaults(),
                java.util.List.of(),
                java.util.List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.SheetLayoutReport(
                "Budget",
                new PaneReport.None(),
                401,
                SheetPresentationReport.defaults(),
                java.util.List.of(),
                java.util.List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindResponse.SheetLayoutReport(
                "Budget",
                new PaneReport.None(),
                100,
                null,
                java.util.List.of(),
                java.util.List.of()));
  }

  @Test
  void sheetPresentationInputsAndReportsNormalizeAndValidate() {
    SheetPresentationInput defaultPresentation =
        new SheetPresentationInput(null, null, null, null, null);
    SheetPresentationInput explicitPresentation =
        new SheetPresentationInput(
            new SheetDisplayInput(false, false, false, true, true),
            new ColorInput("#112233"),
            new SheetOutlineSummaryInput(false, false),
            new SheetDefaultsInput(11, 18.5d),
            java.util.List.of(
                new IgnoredErrorInput(
                    "A1:B2",
                    java.util.List.of(
                        dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType.FORMULA))));

    assertEquals(SheetDisplayInput.defaults(), defaultPresentation.display());
    assertNull(defaultPresentation.tabColor());
    assertEquals(SheetOutlineSummaryInput.defaults(), defaultPresentation.outlineSummary());
    assertEquals(SheetDefaultsInput.defaults(), defaultPresentation.sheetDefaults());
    assertEquals(java.util.List.of(), defaultPresentation.ignoredErrors());
    assertEquals(
        new SheetDisplayInput(false, false, false, true, true), explicitPresentation.display());
    assertEquals(new ColorInput("#112233"), explicitPresentation.tabColor());
    assertEquals(new SheetOutlineSummaryInput(false, false), explicitPresentation.outlineSummary());
    assertEquals(new SheetDefaultsInput(11, 18.5d), explicitPresentation.sheetDefaults());
    assertEquals(
        java.util.List.of(
            new IgnoredErrorInput("A1:B2", java.util.List.of(ExcelIgnoredErrorType.FORMULA))),
        explicitPresentation.ignoredErrors());
    assertEquals(
        new SheetPresentationReport(
            SheetDisplayReport.defaults(),
            null,
            SheetOutlineSummaryReport.defaults(),
            SheetDefaultsReport.defaults(),
            java.util.List.of()),
        SheetPresentationReport.defaults());
    assertEquals(
        new SheetPresentationReport(
            new SheetDisplayReport(false, false, false, true, true),
            new CellColorReport("#112233"),
            new SheetOutlineSummaryReport(false, false),
            new SheetDefaultsReport(11, 18.5d),
            java.util.List.of(
                new IgnoredErrorReport("A1:B2", java.util.List.of(ExcelIgnoredErrorType.FORMULA)))),
        new SheetPresentationReport(
            new SheetDisplayReport(false, false, false, true, true),
            new CellColorReport("#112233"),
            new SheetOutlineSummaryReport(false, false),
            new SheetDefaultsReport(11, 18.5d),
            java.util.List.of(
                new IgnoredErrorReport(
                    "A1:B2", java.util.List.of(ExcelIgnoredErrorType.FORMULA)))));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsInput(0, 15.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new SheetDefaultsInput(ExcelSheetLayoutLimits.MAX_DEFAULT_COLUMN_WIDTH + 1, 15.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new SheetDefaultsInput(8, Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsReport(8, 0.0d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetPresentationInput(
                null,
                null,
                null,
                null,
                java.util.List.of(
                    new IgnoredErrorInput("A1", java.util.List.of(ExcelIgnoredErrorType.FORMULA)),
                    new IgnoredErrorInput(
                        "A1", java.util.List.of(ExcelIgnoredErrorType.EVALUATION_ERROR)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new IgnoredErrorInput(
                "A1",
                java.util.List.of(ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.FORMULA)));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                null,
                SheetOutlineSummaryReport.defaults(),
                SheetDefaultsReport.defaults(),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                null,
                SheetOutlineSummaryReport.defaults(),
                SheetDefaultsReport.defaults(),
                java.util.List.of(
                    new IgnoredErrorReport("A1", java.util.List.of(ExcelIgnoredErrorType.FORMULA)),
                    new IgnoredErrorReport(
                        "A1", java.util.List.of(ExcelIgnoredErrorType.EVALUATION_ERROR)))));
  }

  @Test
  void sheetDisplayDefaultsAndIgnoredErrorDtosCoverConstructorBranches() {
    assertEquals(SheetDisplayInput.defaults(), new SheetDisplayInput(null, null, null, null, null));
    assertEquals(
        new SheetDisplayInput(false, false, true, true, false),
        new SheetDisplayInput(false, false, null, true, null));
    assertEquals(SheetOutlineSummaryInput.defaults(), new SheetOutlineSummaryInput(null, null));
    assertEquals(
        new SheetOutlineSummaryInput(true, false), new SheetOutlineSummaryInput(null, false));
    assertEquals(SheetDefaultsInput.defaults(), new SheetDefaultsInput(null, null));
    assertEquals(new SheetDefaultsInput(8, 18.5d), new SheetDefaultsInput(null, 18.5d));
    assertEquals(SheetDefaultsReport.defaults(), new SheetDefaultsReport(8, 15.0d));

    java.util.List<ExcelIgnoredErrorType> errorTypesWithNull =
        new java.util.ArrayList<>(java.util.Arrays.asList(ExcelIgnoredErrorType.FORMULA, null));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsInput(8, 0.0d));
    assertThrows(
        IllegalArgumentException.class, () -> new SheetDefaultsInput(8, Double.POSITIVE_INFINITY));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsReport(0, 15.0d));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsReport(8, Double.NaN));
    assertThrows(
        NullPointerException.class,
        () -> new IgnoredErrorInput(null, java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new IgnoredErrorInput(" ", java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertThrows(NullPointerException.class, () -> new IgnoredErrorInput("A1", null));
    assertThrows(
        IllegalArgumentException.class, () -> new IgnoredErrorInput("A1", java.util.List.of()));
    assertThrows(NullPointerException.class, () -> new IgnoredErrorInput("A1", errorTypesWithNull));
    assertThrows(
        NullPointerException.class,
        () -> new IgnoredErrorReport(null, java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new IgnoredErrorReport(" ", java.util.List.of(ExcelIgnoredErrorType.FORMULA)));
    assertThrows(NullPointerException.class, () -> new IgnoredErrorReport("A1", null));
    assertThrows(
        IllegalArgumentException.class, () -> new IgnoredErrorReport("A1", java.util.List.of()));
    assertThrows(
        NullPointerException.class, () -> new IgnoredErrorReport("A1", errorTypesWithNull));
  }
}
