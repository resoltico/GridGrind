package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/**
 * Deliberately non-public runtime-type usage that must remain outside the exported surface rule.
 */
public final class ArchitecturePrivateRuntimeLeakFixture {
  private ArchitecturePrivateRuntimeLeakFixture() {}

  static GridGrindRequestDoctor internalRuntimeType(GridGrindRequestDoctor runtimeType) {
    return runtimeType;
  }
}
