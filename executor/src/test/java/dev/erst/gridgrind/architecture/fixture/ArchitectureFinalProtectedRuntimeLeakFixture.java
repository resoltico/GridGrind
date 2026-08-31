package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/** Provides protected implementation usage that is not externally extensible from a final class. */
public final class ArchitectureFinalProtectedRuntimeLeakFixture {
  /** Retains an internal runtime type without exposing it to subclasses. */
  @SuppressWarnings("ProtectedMembersInFinalClass")
  protected GridGrindRequestDoctor internalRuntimeType(GridGrindRequestDoctor runtimeType) {
    return runtimeType;
  }
}
