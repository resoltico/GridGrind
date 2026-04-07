package dev.erst.gridgrind.jazzer.support;

import java.util.Arrays;
import java.util.Objects;

/** Replays GridGrind's scalar fuzz-input surface in pure Java without Jazzer's native layer. */
final class ReplayGridGrindFuzzData implements GridGrindFuzzData {
  private static final double TWO_TO_63 = 0x1.0p63;
  private static final double UNSIGNED_LONG_MAX_AS_DOUBLE = unsignedLongToDouble(-1L);

  private final byte[] data;
  private int remainingBytes;

  ReplayGridGrindFuzzData(byte[] input) {
    Objects.requireNonNull(input, "input must not be null");
    data = Arrays.copyOf(input, input.length);
    remainingBytes = input.length;
  }

  @Override
  public boolean consumeBoolean() {
    return (consumeScalarForRange(1L, 1) & 1L) != 0L;
  }

  @Override
  public byte consumeByte() {
    return (byte) consumeScalarForRange(0xFFL, 1);
  }

  @Override
  public int consumeInt(int min, int max) {
    if (min > max) {
      throw new IllegalArgumentException(
          String.format("min must be <= max (got min: %d, max: %d)", min, max));
    }

    long range = (long) max - min;
    long result = consumeScalarForRange(range, Integer.BYTES);
    int candidate = (int) result;
    if (candidate >= min && candidate <= max) {
      return candidate;
    }
    return min + (int) Long.remainderUnsigned(result, range + 1L);
  }

  @Override
  public double consumeRegularDouble(double min, double max) {
    if (!Double.isFinite(min)) {
      throw new IllegalArgumentException("min must be a regular double");
    }
    if (!Double.isFinite(max)) {
      throw new IllegalArgumentException("max must be a regular double");
    }
    if (min > max) {
      throw new IllegalArgumentException(
          String.format("min must be <= max (got min: %f, max: %f)", min, max));
    }

    double range;
    double result = min;
    if (min < 0.0d && max > 0.0d && min + Double.MAX_VALUE < max) {
      range = (max / 2.0d) - (min / 2.0d);
      if (consumeBoolean()) {
        result += range;
      }
    } else {
      range = max - min;
    }

    result += range * consumeProbabilityDouble();
    return Math.min(result, max);
  }

  @Override
  public int remainingBytes() {
    return remainingBytes;
  }

  private long consumeScalarForRange(long range, int maxBytes) {
    long result = 0L;
    int offset = 0;
    while (offset < Byte.SIZE * maxBytes && (range >>> offset) > 0L && remainingBytes != 0) {
      remainingBytes--;
      result = (result << Byte.SIZE) | Byte.toUnsignedLong(data[remainingBytes]);
      offset += Byte.SIZE;
    }
    return result;
  }

  private long consumeUnsignedLong() {
    long result = 0L;
    int consumedBytes = 0;
    while (consumedBytes < Long.BYTES && remainingBytes != 0) {
      remainingBytes--;
      result = (result << Byte.SIZE) | Byte.toUnsignedLong(data[remainingBytes]);
      consumedBytes++;
    }
    return result;
  }

  private double consumeProbabilityDouble() {
    return unsignedLongToDouble(consumeUnsignedLong()) / UNSIGNED_LONG_MAX_AS_DOUBLE;
  }

  private static double unsignedLongToDouble(long value) {
    return value >= 0L ? (double) value : (double) (value & Long.MAX_VALUE) + TWO_TO_63;
  }
}
