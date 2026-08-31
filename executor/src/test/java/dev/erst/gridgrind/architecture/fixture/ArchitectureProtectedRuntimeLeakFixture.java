package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/** Deliberately extensible API shape with a protected runtime-type leak. */
public class ArchitectureProtectedRuntimeLeakFixture {
  /** Exposes a runtime type to subclasses solely for architecture-rule regression. */
  protected GridGrindRequestDoctor leakedRuntimeType() {
    throw new UnsupportedOperationException("fixture method must not execute");
  }
}
