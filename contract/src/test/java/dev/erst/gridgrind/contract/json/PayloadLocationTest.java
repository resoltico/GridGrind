package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for the serializable payload-location variants used by payload exceptions. */
class PayloadLocationTest {
  @Test
  void fromBuildsEverySupportedLocationShapeAndCollapsesPartialCursorInputs() {
    PayloadLocation unavailable =
        PayloadLocation.from(Optional.empty(), Optional.empty(), Optional.empty());
    PayloadLocation pathOnly =
        PayloadLocation.from(Optional.of("steps[0].target"), Optional.empty(), Optional.empty());
    PayloadLocation lineColumn =
        PayloadLocation.from(Optional.empty(), Optional.of(9), Optional.of(4));
    PayloadLocation located =
        PayloadLocation.from(Optional.of("steps[0].target"), Optional.of(9), Optional.of(4));
    PayloadLocation collapsedToPathOnly =
        PayloadLocation.from(Optional.of("steps[0].target"), Optional.of(9), Optional.empty());
    PayloadLocation collapsedToUnavailable =
        PayloadLocation.from(Optional.empty(), Optional.empty(), Optional.of(4));

    assertInstanceOf(PayloadLocation.Unavailable.class, unavailable);
    assertEquals(Optional.empty(), unavailable.jsonPath());
    assertEquals(Optional.empty(), unavailable.jsonLine());
    assertEquals(Optional.empty(), unavailable.jsonColumn());

    assertInstanceOf(PayloadLocation.PathOnly.class, pathOnly);
    assertEquals(Optional.of("steps[0].target"), pathOnly.jsonPath());
    assertEquals(Optional.empty(), pathOnly.jsonLine());
    assertEquals(Optional.empty(), pathOnly.jsonColumn());

    assertInstanceOf(PayloadLocation.LineColumn.class, lineColumn);
    assertEquals(Optional.empty(), lineColumn.jsonPath());
    assertEquals(Optional.of(9), lineColumn.jsonLine());
    assertEquals(Optional.of(4), lineColumn.jsonColumn());

    assertInstanceOf(PayloadLocation.Located.class, located);
    assertEquals(Optional.of("steps[0].target"), located.jsonPath());
    assertEquals(Optional.of(9), located.jsonLine());
    assertEquals(Optional.of(4), located.jsonColumn());

    assertInstanceOf(PayloadLocation.PathOnly.class, collapsedToPathOnly);
    assertEquals(Optional.of("steps[0].target"), collapsedToPathOnly.jsonPath());
    assertEquals(Optional.empty(), collapsedToPathOnly.jsonLine());
    assertEquals(Optional.empty(), collapsedToPathOnly.jsonColumn());

    assertInstanceOf(PayloadLocation.Unavailable.class, collapsedToUnavailable);
    assertEquals(Optional.empty(), collapsedToUnavailable.jsonPath());
    assertEquals(Optional.empty(), collapsedToUnavailable.jsonLine());
    assertEquals(Optional.empty(), collapsedToUnavailable.jsonColumn());
  }

  @Test
  void variantsRejectBlankPathsAndNonPositiveCursorValues() {
    assertThrows(IllegalArgumentException.class, () -> new PayloadLocation.PathOnly(" "));
    assertThrows(IllegalArgumentException.class, () -> new PayloadLocation.LineColumn(0, 4));
    assertThrows(IllegalArgumentException.class, () -> new PayloadLocation.LineColumn(9, 0));
    assertThrows(IllegalArgumentException.class, () -> new PayloadLocation.Located(" ", 9, 4));
    assertThrows(
        IllegalArgumentException.class, () -> new PayloadLocation.Located("steps[0]", 0, 4));
    assertThrows(
        IllegalArgumentException.class, () -> new PayloadLocation.Located("steps[0]", 9, 0));
  }
}
