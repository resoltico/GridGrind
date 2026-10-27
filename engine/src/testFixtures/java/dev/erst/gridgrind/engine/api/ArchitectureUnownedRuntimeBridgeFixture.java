package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/** Deliberately wires an unapproved API class to the internal runtime. */
public final class ArchitectureUnownedRuntimeBridgeFixture {
  private ArchitectureUnownedRuntimeBridgeFixture() {}

  /** Introduces an internal dependency without exposing it in the method's declared type. */
  public static Class<?> runtimeType() {
    return GridGrindRequestDoctor.class;
  }
}
