package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class ReplayGridGrindFuzzDataTest {
  @Test
  void replayConsumesScalarBytesFromTheInputTail() {
    GridGrindFuzzData data = GridGrindFuzzData.replay(new byte[] {1, 2, 3});

    assertEquals(3, Byte.toUnsignedInt(data.consumeByte()));
    assertEquals(2, data.remainingBytes());
    assertEquals(2, Byte.toUnsignedInt(data.consumeByte()));
    assertEquals(1, data.remainingBytes());
    assertTrue(data.consumeBoolean());
    assertEquals(0, data.remainingBytes());
  }

  @Test
  void replayConsumesOnlyTheBytesNeededForIntRanges() {
    GridGrindFuzzData data = GridGrindFuzzData.replay(new byte[] {1, 2, 3});

    assertEquals(0x0302, data.consumeInt(0, 0xFFFF));
    assertEquals(1, data.remainingBytes());
  }

  @Test
  void replayAppliesJazzerStyleModuloForOutOfRangeInts() {
    GridGrindFuzzData data = GridGrindFuzzData.replay(new byte[] {(byte) 0xFF});

    assertEquals(12, data.consumeInt(10, 20));
    assertEquals(0, data.remainingBytes());
  }

  @Test
  void replayClampsRegularDoubleToTheUpperBound() {
    GridGrindFuzzData data =
        GridGrindFuzzData.replay(
            new byte[] {(byte) 0xFF, (byte) 0xFF, (byte) 0xFF, (byte) 0xFF,
              (byte) 0xFF, (byte) 0xFF, (byte) 0xFF, (byte) 0xFF});

    assertEquals(10.0d, data.consumeRegularDouble(0.0d, 10.0d));
    assertEquals(0, data.remainingBytes());
  }

  @Test
  void replayReturnsDeterministicFallbacksForEmptyInput() {
    GridGrindFuzzData data = GridGrindFuzzData.replay(new byte[0]);

    assertFalse(data.consumeBoolean());
    assertEquals(0, Byte.toUnsignedInt(data.consumeByte()));
    assertEquals(1, data.consumeInt(1, 10));
    assertEquals(-5.0d, data.consumeRegularDouble(-5.0d, 5.0d));
    assertEquals(0, data.remainingBytes());
  }
}
