package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import org.junit.jupiter.api.Test;

/** Verifies an unannotated record component retains its Java field name. */
class InputResolutionOriginsDefaultJsonPropertyTest {
  @Test
  void retainsTheRecordComponentNameWithoutAJsonProperty() {
    assertEquals(
        "text",
        InputResolutionOrigins.jsonFieldName(
            TextSourceInput.Inline.class.getRecordComponents()[0]));
  }
}
