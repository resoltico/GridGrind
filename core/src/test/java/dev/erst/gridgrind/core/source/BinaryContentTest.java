package dev.erst.gridgrind.core.source;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Checks immutable byte ownership and factual empty-payload preservation. */
class BinaryContentTest {
  @Test
  void preservesValueEqualityWithoutExposingMutableBuffers() {
    byte[] input = {1, 2, 3};
    BinaryContent data = new BinaryContent(input);
    input[0] = 7;
    assertArrayEquals(new byte[] {1, 2, 3}, data.bytes());
    byte[] extractedBytes = data.bytes();
    extractedBytes[0] = 9;
    assertArrayEquals(new byte[] {1, 2, 3}, data.bytes());
    assertArrayEquals(new byte[] {1, 2, 3}, data.bytes());
    assertEquals(3, data.size());
    assertEquals(new BinaryContent(new byte[] {1, 2, 3}), data);
    assertEquals(new BinaryContent(new byte[] {1, 2, 3}).hashCode(), data.hashCode());
    assertNotEquals(data, new BinaryContent(new byte[] {3, 2, 1}));
    assertNotEquals(data, "payload");
    assertTrue(data.toString().contains("size=3"));
    assertEquals(0, BinaryContent.readback(new byte[0]).size());
    assertThrows(NullPointerException.class, () -> new BinaryContent(null));
    assertThrows(IllegalArgumentException.class, () -> new BinaryContent(new byte[0]));
  }
}
