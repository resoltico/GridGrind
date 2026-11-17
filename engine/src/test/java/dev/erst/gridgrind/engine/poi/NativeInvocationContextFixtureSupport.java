package dev.erst.gridgrind.engine.poi;

import java.util.List;

/** Supplies injected teardown only to owner-finalization tests across the adapter boundary. */
public final class NativeInvocationContextFixtureSupport {
  private NativeInvocationContextFixtureSupport() {}

  /** Creates an explicit context whose teardown remains owned by the executor under test. */
  public static NativeInvocationContext withTeardown(List<Runnable> teardown) {
    return new NativeInvocationContext(teardown);
  }
}
