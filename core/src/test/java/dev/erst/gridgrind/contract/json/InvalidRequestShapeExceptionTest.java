package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for {@link InvalidRequestShapeException}. */
class InvalidRequestShapeExceptionTest {
  @Test
  void exposesRecordedJsonLocationMetadata() {
    RuntimeException cause = new RuntimeException("boom");
    InvalidRequestShapeException exception =
        new InvalidRequestShapeException(
            new MessageShape("bad shape", Optional.of("reads[0]")),
            Optional.of("reads[0]"),
            Optional.of(4),
            Optional.of(12),
            cause);

    assertEquals("bad shape", exception.getMessage());
    assertEquals(cause, exception.getCause());
    assertEquals("reads[0]", exception.jsonPath().orElseThrow());
    assertEquals(4, exception.jsonLine().orElseThrow());
    assertEquals(12, exception.jsonColumn().orElseThrow());
  }

  @Test
  void allowsMissingJsonLocationMetadata() {
    InvalidRequestShapeException exception =
        new InvalidRequestShapeException(
            new MessageShape("bad shape", Optional.empty()),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            null);

    assertTrue(exception.jsonPath().isEmpty());
    assertTrue(exception.jsonLine().isEmpty());
    assertTrue(exception.jsonColumn().isEmpty());
  }
}
