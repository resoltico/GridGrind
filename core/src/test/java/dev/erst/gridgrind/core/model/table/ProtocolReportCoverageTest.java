package dev.erst.gridgrind.core.model.table;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.layout.SheetDefaultsReport;
import dev.erst.gridgrind.core.model.layout.SheetDisplayReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorReport;
import dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Residual coverage for factual report records and explicit autofilter sort variants. */
class ProtocolReportCoverageTest {
  @Test
  void autofilterSortConditionVariantsValidateRangesAndColors() {
    AutofilterSortConditionSpec.FontColor authoredFontColor =
        new AutofilterSortConditionSpec.FontColor(
            new RangeRef("A1:A3"), true, ColorSpec.theme(4, Double.valueOf(0.15d)));
    AutofilterSortConditionReport.FontColor reportedFontColor =
        new AutofilterSortConditionReport.FontColor(
            new RangeLabel("A1:A3"), false, CellColorReport.rgb(new RgbHex("#AABBCC")));

    assertEquals("A1:A3", authoredFontColor.range().value());
    assertFalse(reportedFontColor.descending());
    assertEquals(CellColorReport.rgb(new RgbHex("#AABBCC")), reportedFontColor.color());
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionReport.Value(new RangeLabel(" "), false));
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
        new IgnoredErrorReport(
            new RangeLabel("A1:A2"), List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT));
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
        new AutofilterSortConditionReport.FontColor(
            new RangeLabel("A1:A3"), false, CellColorReport.rgb(new RgbHex("#AABBCC")));

    AutofilterSortStateReport normalized =
        new AutofilterSortStateReport(
            new RangeLabel("A1:A3"), true, false, Optional.empty(), List.of(reportedFontColor));

    assertEquals(Optional.empty(), normalized.sortMethod());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterSortStateReport(
                new RangeLabel(" "),
                false,
                false,
                Optional.of(ExcelAutofilterSortMethod.STROKE),
                List.of()));
  }
}
