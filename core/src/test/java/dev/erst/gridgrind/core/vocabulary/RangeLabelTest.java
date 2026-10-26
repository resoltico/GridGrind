package dev.erst.gridgrind.core.vocabulary;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Factual range text can expose corruption without relaxing authored coordinates. */
class RangeLabelTest {
  @Test
  void retainsBlankAndMalformedStoredText() {
    assertEquals("", new RangeLabel("").value());
    assertEquals(" ", new RangeLabel(" ").value());
    assertEquals("broken range", new RangeLabel("broken range").value());
    assertThrows(NullPointerException.class, () -> new RangeLabel(null));
  }

  @Test
  void authoredRangesStillRejectBlankAndMalformedText() {
    assertThrows(IllegalArgumentException.class, () -> new RangeRef(""));
    assertThrows(IllegalArgumentException.class, () -> new RangeRef(" "));
    assertThrows(IllegalArgumentException.class, () -> new RangeRef("broken range"));
  }
}
