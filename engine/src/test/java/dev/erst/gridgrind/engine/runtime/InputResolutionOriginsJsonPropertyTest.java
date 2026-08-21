package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import com.fasterxml.jackson.annotation.JsonProperty;
import org.junit.jupiter.api.Test;

/** Verifies explicit non-blank JSON property names win over Java record-component names. */
class InputResolutionOriginsJsonPropertyTest {
  @Test
  void usesAnExplicitJsonPropertyName() {
    assertEquals(
        "wireName",
        InputResolutionOrigins.jsonFieldName(RenamedSource.class.getRecordComponents()[0]));
  }

  private record RenamedSource(@JsonProperty("wireName") String source) {}
}
