package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;
import tools.jackson.databind.DatabindException;
import tools.jackson.databind.json.JsonMapper;

/** Serializer coverage for selector runtime encoding and reflective access failures. */
class SelectorJsonSerializerTest {
  @Test
  void serializesCanonicalRecordSelectorsWithTypeAndFields() throws Exception {
    String json =
        JsonMapper.builder()
            .build()
            .writeValueAsString(new SelectorEnvelope(new CellSelector.ByAddress("Budget", "A1")));

    assertTrue(json.contains("\"selector\":{\"type\":\"BY_ADDRESS\""));
    assertTrue(json.contains("\"sheetName\":\"Budget\""));
    assertTrue(json.contains("\"address\":\"A1\""));
  }

  @Test
  void rejectsNonRecordSelectorsAndMismatchedRecordAccessors() throws Exception {
    DatabindException nonRecordFailure =
        assertThrows(
            DatabindException.class,
            () ->
                JsonMapper.builder()
                    .build()
                    .writeValueAsString(new SelectorEnvelope(new NonRecordSelector())));
    assertTrue(nonRecordFailure.getMessage().contains("Selector runtime type must be a record"));
    assertEquals(
        "Selector runtime type must be a record: " + NonRecordSelector.class,
        nonRecordFailure.getCause().getMessage());

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

  /** Non-record selector used to verify serializer rejection of invalid runtime types. */
  private static final class NonRecordSelector implements Selector {
    @Override
    public SelectorCardinality cardinality() {
      return SelectorCardinality.EXACTLY_ONE;
    }
  }

  private record SelectorEnvelope(Selector selector) {}
}
