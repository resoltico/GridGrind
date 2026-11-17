package dev.erst.gridgrind.core.source;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Basic-decoder source grammar remains exact while syntax validation retains no payload array. */
class BinarySourceEncodingTest {
  @Test
  void paddedAndUnpaddedSourcesHaveExactDecodedLengths() {
    assertEquals(1, BinarySourceSpec.decodedSize(new BinarySourceSpec.InlineBase64("AA==")));
    assertEquals(1, BinarySourceSpec.decodedSize(new BinarySourceSpec.InlineBase64("AB")));
    assertEquals(2, BinarySourceSpec.decodedSize(new BinarySourceSpec.InlineBase64("AAE=")));
    assertEquals(2, BinarySourceSpec.decodedSize(new BinarySourceSpec.InlineBase64("AAF")));
    assertEquals(3, BinarySourceSpec.decodedSize(new BinarySourceSpec.InlineBase64("AAEC")));
  }

  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void paddingNeverHidesTrailingInputAndMimeWhitespaceIsNotAccepted() {
    for (String value :
        List.of(
            "A",
            "AA=",
            "AA===",
            "AA==A",
            "AA==AAAA",
            "AA==\n",
            "AAE=A",
            "AA EC",
            "AA\u00ff",
            "_w==",
            "/w==!")) {
      assertThrows(
          IllegalArgumentException.class, () -> new BinarySourceSpec.InlineBase64(value), value);
    }
  }
}
