package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.lang.reflect.InvocationTargetException;
import org.junit.jupiter.api.Test;

/** Verifies reflective request traversal does not silently hide an accessor failure. */
class InputResolutionOriginsAccessorFailureTest {
  @Test
  void turnsAnAccessorFailureIntoOneInvariantFailure() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                InputResolutionOrigins.componentValue(
                    new BrokenAccessor("value"), BrokenAccessor.class.getRecordComponents()[0]));

    assertEquals("Unable to inspect request component source", failure.getMessage());
    assertInstanceOf(InvocationTargetException.class, failure.getCause());
  }

  private record BrokenAccessor(String source) {
    @Override
    @SuppressWarnings("UnusedMethod") // Invoked by the record-component reflection under test.
    public String source() {
      throw new IllegalStateException("accessor failure");
    }
  }
}
