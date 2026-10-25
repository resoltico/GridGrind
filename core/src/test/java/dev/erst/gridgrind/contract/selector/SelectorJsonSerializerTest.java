package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;
import tools.jackson.databind.json.JsonMapper;

/** Serializer coverage for selector runtime encoding and reflective access failures. */
class SelectorJsonSerializerTest {
  @Test
  void serializesCanonicalRecordSelectorsWithTypeAndFields() throws Exception {
    String json =
        JsonMapper.builder()
            .build()
            .writeValueAsString(new SelectorEnvelope(new CellSelector.ByAddress("Budget", "A1")));

    assertTrue(json.contains("\"selector\":{\"type\":\"CELL_BY_ADDRESS\""));
    assertTrue(json.contains("\"sheetName\":\"Budget\""));
    assertTrue(json.contains("\"address\":\"A1\""));
  }

  @Test
  void selectorRootIsClosedAndMismatchedRecordAccessorsFailFast() {
    assertTrue(Selector.class.isSealed());
    assertEquals(
        java.util.Set.of(
            WorkbookSelector.class,
            SheetSelector.class,
            CellSelector.class,
            RangeSelector.class,
            RowBandSelector.class,
            ColumnBandSelector.class,
            DrawingObjectSelector.class,
            ChartSelector.class,
            TableSelector.class,
            PivotTableSelector.class,
            NamedRangeSelector.class,
            TableRowSelector.class,
            TableCellSelector.class),
        java.util.Set.of(Selector.class.getPermittedSubclasses()));

    java.lang.reflect.RecordComponent component =
        CellSelector.ByAddress.class.getRecordComponents()[0];

    IllegalStateException invocationTargetException =
        assertThrows(
            IllegalStateException.class,
            () ->
                SelectorJsonSerializer.readComponent(
                    new SheetSelector.ByName("Budget"), component));
    assertEquals(
        "Unable to read selector component sheetName from " + SheetSelector.ByName.class.getName(),
        invocationTargetException.getMessage());
  }

  private record SelectorEnvelope(Selector selector) {}
}
