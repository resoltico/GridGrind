package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/** Provides protected implementation usage that is not externally extensible from a final class. */
public final class ArchitectureFinalProtectedRuntimeLeakFixture {
  /** Deliberately retains a protected member to prove final enclosing types do not leak it. */
  @SuppressWarnings({"PMD.ProtectedMemberInFinalClass", "ProtectedMembersInFinalClass"})
  protected GridGrindRequestDoctor internalRuntimeType(GridGrindRequestDoctor runtimeType) {
    return runtimeType;
  }
}
