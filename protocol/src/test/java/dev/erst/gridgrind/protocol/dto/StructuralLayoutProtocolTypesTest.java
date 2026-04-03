package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Direct tests for the structural-layout protocol input and report families. */
class StructuralLayoutProtocolTypesTest {
  @Test
  void paneInputVariantsValidateTheirContracts() {
    assertEquals(new PaneInput.None(), new PaneInput.None());
    assertEquals(new PaneInput.Frozen(1, 2, 1, 2), new PaneInput.Frozen(1, 2, 1, 2));
    assertEquals(
        new PaneInput.Split(0, 1800, 0, 3, PaneRegion.LOWER_LEFT),
        new PaneInput.Split(0, 1800, 0, 3, PaneRegion.LOWER_LEFT));

    assertThrows(IllegalArgumentException.class, () -> new PaneInput.Frozen(-1, 2, 1, 2));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneInput.Split(0, 0, 0, 0, PaneRegion.UPPER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneInput.Split(0, 1800, 1, 3, PaneRegion.LOWER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneInput.Split(1200, 0, 2, 1, PaneRegion.UPPER_RIGHT));
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
    assertEquals(PrintOrientation.PORTRAIT, input.orientation());
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
        new PaneReport.Split(1200, 2400, 3, 4, PaneRegion.LOWER_RIGHT),
        new PaneReport.Split(1200, 2400, 3, 4, PaneRegion.LOWER_RIGHT));
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
            PrintOrientation.PORTRAIT,
            new PrintScalingReport.Automatic(),
            new PrintTitleRowsReport.None(),
            new PrintTitleColumnsReport.None(),
            new HeaderFooterTextReport("", "", ""),
            new HeaderFooterTextReport("", "", "")),
        new PrintLayoutReport(
            "Budget",
            new PrintAreaReport.None(),
            PrintOrientation.PORTRAIT,
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
                PrintOrientation.PORTRAIT,
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
                "Budget", null, 100, java.util.List.of(), java.util.List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.SheetLayoutReport(
                "Budget", new PaneReport.None(), 9, java.util.List.of(), java.util.List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.SheetLayoutReport(
                "Budget", new PaneReport.None(), 401, java.util.List.of(), java.util.List.of()));
  }
}
