package dev.erst.gridgrind.fuzz.support;

/** Exposes the narrow scalar fuzz-input surface used by GridGrind's structured generators. */
public interface GridGrindFuzzData {
  /** Creates a pure-Java replay cursor over committed raw fuzz-input bytes. */
  static GridGrindFuzzData replay(byte[] input) {
    return new ReplayGridGrindFuzzData(input);
  }

  /** Consumes one boolean value from the current fuzz-input cursor. */
  boolean consumeBoolean();

  /** Consumes one byte value from the current fuzz-input cursor. */
  byte consumeByte();

  /** Consumes one integer within the supplied inclusive bounds. */
  int consumeInt(int min, int max);

  /** Consumes one regular finite double within the supplied inclusive bounds. */
  double consumeRegularDouble(double min, double max);

  /** Returns the number of raw bytes that remain unconsumed. */
  int remainingBytes();
}
