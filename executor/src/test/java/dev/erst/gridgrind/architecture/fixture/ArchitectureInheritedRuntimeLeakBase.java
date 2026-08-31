package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/** Deliberately invalid base type that exposes an unexported runtime type. */
public class ArchitectureInheritedRuntimeLeakBase {
  /** Exposes a runtime type solely for architecture-rule regression. */
  public GridGrindRequestDoctor leakedRuntimeType() {
    throw new UnsupportedOperationException("fixture method must not execute");
  }
}
