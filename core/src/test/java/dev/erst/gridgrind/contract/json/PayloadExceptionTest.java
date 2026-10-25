package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Constructor guard tests for payload exceptions with structured JSON location metadata. */
class PayloadExceptionTest {
  @Test
  void payloadExceptionsRequireExplicitOptionalLocationContainers() {
    assertEquals(
        "jsonPath must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new InvalidJsonException(
                        "bad json", null, Optional.empty(), Optional.empty(), null))
            .getMessage());
    assertEquals(
        "jsonLine must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new InvalidRequestException(
                        new MessageInvariant("bad request", Optional.empty()),
                        Optional.empty(),
                        null,
                        Optional.empty(),
                        null))
            .getMessage());
    assertEquals(
        "jsonColumn must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new InvalidRequestShapeException(
                        new MessageShape("bad shape", Optional.empty()),
                        Optional.empty(),
                        Optional.empty(),
                        null,
                        null))
            .getMessage());
  }

  @Test
  void payloadExceptionsExposeNormalizedTypedLocations() {
    PayloadException encoding =
        new InvalidEncodingException(
            "bad encoding", Optional.empty(), Optional.empty(), Optional.empty(), null);
    PayloadException pathOnly =
        new InvalidJsonException(
            "bad json", Optional.of("steps[0].target"), Optional.of(4), Optional.empty(), null);
    PayloadException lineColumn =
        new InvalidRequestException(
            new MessageInvariant("bad request", Optional.empty()),
            Optional.empty(),
            Optional.of(7),
            Optional.of(3),
            null);
    PayloadException unavailable =
        new InvalidRequestShapeException(
            new MessageShape("bad shape", Optional.empty()),
            Optional.empty(),
            Optional.empty(),
            Optional.of(3),
            null);

    assertInstanceOf(PayloadLocation.Unavailable.class, encoding.jsonLocation());
    assertInstanceOf(PayloadLocation.PathOnly.class, pathOnly.jsonLocation());
    assertEquals(Optional.of("steps[0].target"), pathOnly.jsonPath());
    assertEquals(Optional.empty(), pathOnly.jsonLine());
    assertEquals(Optional.empty(), pathOnly.jsonColumn());

    assertInstanceOf(PayloadLocation.LineColumn.class, lineColumn.jsonLocation());
    assertEquals(Optional.empty(), lineColumn.jsonPath());
    assertEquals(Optional.of(7), lineColumn.jsonLine());
    assertEquals(Optional.of(3), lineColumn.jsonColumn());

    assertInstanceOf(PayloadLocation.Unavailable.class, unavailable.jsonLocation());
    assertEquals(Optional.empty(), unavailable.jsonPath());
    assertEquals(Optional.empty(), unavailable.jsonLine());
    assertEquals(Optional.empty(), unavailable.jsonColumn());
  }
}
