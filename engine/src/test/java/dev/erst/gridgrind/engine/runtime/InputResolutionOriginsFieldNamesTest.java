package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.source.TextSourceSpec;
import org.junit.jupiter.api.Test;

/** Verifies canonical record component names identify resolved source origins. */
class InputResolutionOriginsFieldNamesTest {
  @Test
  void usesCanonicalRecordComponentNames() {
    assertEquals(
        "text",
        InputResolutionOrigins.jsonFieldName(TextSourceSpec.Inline.class.getRecordComponents()[0]));
  }
}
