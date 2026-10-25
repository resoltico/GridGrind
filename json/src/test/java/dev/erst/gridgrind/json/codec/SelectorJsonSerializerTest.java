package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import org.junit.jupiter.api.Test;

/** Serializer coverage for selector runtime encoding and reflective access failures. */
class SelectorJsonSerializerTest {
  @Test
  void serializesCanonicalRecordSelectorsWithTypeAndFields() throws Exception {
    String json =
        GridGrindJsonMapperSupport.JSON_MAPPER.writeValueAsString(
            new SelectorEnvelope(new CellSelector.ByAddress(new SheetName("Budget"), "A1")));

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
                    new SheetSelector.ByName(new SheetName("Budget")), component));
    assertEquals(
        "Unable to read selector component sheetName from " + SheetSelector.ByName.class.getName(),
        invocationTargetException.getMessage());
  }

  private record SelectorEnvelope(Selector selector) {}
}
