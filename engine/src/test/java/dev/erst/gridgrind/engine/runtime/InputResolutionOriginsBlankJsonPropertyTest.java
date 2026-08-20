package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import com.fasterxml.jackson.annotation.JsonProperty;
import org.junit.jupiter.api.Test;

/** Verifies a blank JSON-property annotation retains the Java record-component name. */
class InputResolutionOriginsBlankJsonPropertyTest {
  @Test
  void retainsTheRecordComponentNameForABlankJsonProperty() {
    assertEquals(
        "source",
        InputResolutionOrigins.jsonFieldName(DefaultNamedSource.class.getRecordComponents()[0]));
  }

  private record DefaultNamedSource(@JsonProperty String source) {}
}
