package dev.erst.gridgrind.protocol.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import org.junit.jupiter.api.Test;

/** Tests for {@link InvalidRequestShapeException}. */
class InvalidRequestShapeExceptionTest {
  @Test
  void exposesRecordedJsonLocationMetadata() {
    RuntimeException cause = new RuntimeException("boom");
    InvalidRequestShapeException exception =
        new InvalidRequestShapeException("bad shape", "reads[0]", 4, 12, cause);

    assertEquals("bad shape", exception.getMessage());
    assertEquals(cause, exception.getCause());
    assertEquals("reads[0]", exception.jsonPath());
    assertEquals(4, exception.jsonLine());
    assertEquals(12, exception.jsonColumn());
  }

  @Test
  void allowsMissingJsonLocationMetadata() {
    InvalidRequestShapeException exception =
        new InvalidRequestShapeException("bad shape", null, null, null, null);

    assertNull(exception.jsonPath());
    assertNull(exception.jsonLine());
    assertNull(exception.jsonColumn());
  }
}
