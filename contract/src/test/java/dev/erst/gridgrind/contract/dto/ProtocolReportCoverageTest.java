package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Residual coverage for factual report records and explicit autofilter sort variants. */
class ProtocolReportCoverageTest {
  @Test
  void autofilterSortConditionVariantsValidateRangesAndColors() {
    AutofilterSortConditionInput.FontColor authoredFontColor =
        new AutofilterSortConditionInput.FontColor(
            "A1:A3", true, ColorInput.theme(4, Double.valueOf(0.15d)));
    AutofilterSortConditionReport.FontColor reportedFontColor =
        new AutofilterSortConditionReport.FontColor("A1:A3", false, CellColorReport.rgb("#AABBCC"));

    assertEquals("A1:A3", authoredFontColor.range());
    assertFalse(reportedFontColor.descending());
    assertEquals(CellColorReport.rgb("#AABBCC"), reportedFontColor.color());
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterSortConditionReport.Value(" ", false));
  }

  @Test
  void reportConstructorsNormalizeOptionalTextAndRejectDuplicateIgnoredErrorRanges() {
    TableColumnReport normalized =
        new TableColumnReport(
            7L, "Amount", Optional.of(" "), Optional.of(" "), Optional.of("sum"), Optional.of(" "));

    assertEquals(Optional.empty(), normalized.uniqueName());
    assertEquals(Optional.empty(), normalized.totalsRowLabel());
    assertEquals(Optional.of("sum"), normalized.totalsRowFunction());
    assertEquals(Optional.empty(), normalized.calculatedColumnFormula());

    IgnoredErrorReport ignored =
        new IgnoredErrorReport("A1:A2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                Optional.empty(),
                SheetOutlineSummaryReport.defaults(),
                SheetDefaultsReport.defaults(),
                List.of(ignored, ignored)));
  }

  @Test
  void sortStateReportCreatorCoversBlankValidationAndExplicitEmptySortMethod() {
    AutofilterSortConditionReport.FontColor reportedFontColor =
        new AutofilterSortConditionReport.FontColor("A1:A3", false, CellColorReport.rgb("#AABBCC"));

    AutofilterSortStateReport normalized =
        AutofilterSortStateReport.create(
            "A1:A3", true, false, Optional.empty(), List.of(reportedFontColor));

    assertEquals(Optional.empty(), normalized.sortMethod());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterSortStateReport(
                " ", false, false, Optional.of(ExcelAutofilterSortMethod.STROKE), List.of()));
  }
}
