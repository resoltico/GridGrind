package dev.erst.gridgrind.jazzer.support;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import java.util.Objects;

/** Delegates GridGrind's narrow fuzz-input surface to Jazzer during active fuzzing. */
record JazzerGridGrindFuzzData(FuzzedDataProvider delegate) implements GridGrindFuzzData {
  JazzerGridGrindFuzzData {
    Objects.requireNonNull(delegate, "delegate must not be null");
  }

  @Override
  public boolean consumeBoolean() {
    return delegate.consumeBoolean();
  }

  @Override
  public byte consumeByte() {
    return delegate.consumeByte();
  }

  @Override
  public int consumeInt(int min, int max) {
    return delegate.consumeInt(min, max);
  }

  @Override
  public double consumeRegularDouble(double min, double max) {
    return delegate.consumeRegularDouble(min, max);
  }

  @Override
  public int remainingBytes() {
    return delegate.remainingBytes();
  }
}
