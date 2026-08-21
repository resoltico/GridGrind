package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.source.BinarySourceInput;
import org.junit.jupiter.api.Test;

/** Verifies source-origin field selection for binary transport-only inputs. */
class InputResolutionOriginsTest {
  @Test
  void mapsBinaryStandardInputToItsDiscriminatorField() {
    assertEquals("type", InputResolutionOrigins.sourceField(new BinarySourceInput.StandardInput()));
  }
}
