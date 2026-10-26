package dev.erst.gridgrind.core.model.validation;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing data-validation read reports. */
class DataValidationEntryReportTest {
  @Test
  void supportedAndUnsupportedEntriesCopyRangesAndValidateFields() {
    List<RangeLabel> ranges = new ArrayList<>(List.of(new RangeLabel("A1:A3")));

    DataValidationEntryReport.Supported supported =
        new DataValidationEntryReport.Supported(
            ranges,
            new DataValidationEntryReport.DataValidationDefinitionReport(
                new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
                true,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.Unsupported unsupported =
        new DataValidationEntryReport.Unsupported(
            List.of(new RangeLabel("B1:B3")), "ANY", "Not modeled");
    ranges.clear();

    assertEquals(
        List.of("A1:A3"),
        supported.ranges().stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
    assertEquals("ANY", unsupported.kind());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationEntryReport.Supported(
                List.of(),
                new DataValidationEntryReport.DataValidationDefinitionReport(
                    new DataValidationRuleSpec.ExplicitList(List.of("Queued")),
                    false,
                    false,
                    Optional.empty(),
                    Optional.empty())));
    assertDoesNotThrow(
        () ->
            new DataValidationEntryReport.Unsupported(
                List.of(new RangeLabel(" ")), "ANY", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationEntryReport.Unsupported(
                List.of(new RangeLabel("A1")), " ", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationEntryReport.Unsupported(List.of(new RangeLabel("A1")), "ANY", " "));
  }
}
