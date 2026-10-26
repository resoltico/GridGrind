package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.vocabulary.DefinedNameLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import org.junit.jupiter.api.Test;

/** Observed name facts must not be projected through authoring validation during encoding. */
class NamedRangeObservationCodecTest {
  @Test
  void roundTripsEmptyFormulaAndStoredSheetLabel() {
    var expected =
        new NamedRangeReport.FormulaReport(
            new DefinedNameLabel("EmptyReference"),
            new NamedRangeScopeReport.Sheet(new SheetLabel("stored[]")),
            "");
    var mapper = GridGrindJsonMapperSupport.WIRE_JSON_MAPPER;
    String wire = mapper.writerFor(NamedRangeReport.class).writeValueAsString(expected);
    var tree = mapper.readTree(wire);
    assertEquals("stored[]", tree.path("scope").path("sheetName").asText());
    assertEquals("", tree.path("refersToFormula").asText());
    assertEquals(expected, mapper.readValue(wire, NamedRangeReport.class));
  }
}
