package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/** Provides package-private implementation usage in a type that remains externally extensible. */
public class ArchitectureExtensiblePackageRuntimeLeakFixture {
  GridGrindRequestDoctor internalRuntimeType(GridGrindRequestDoctor runtimeType) {
    return runtimeType;
  }
}
