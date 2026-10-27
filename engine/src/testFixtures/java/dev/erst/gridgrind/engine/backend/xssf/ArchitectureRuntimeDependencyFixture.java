package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;

/** Deliberately introduces a runtime dependency into workbook mechanics. */
public final class ArchitectureRuntimeDependencyFixture {
  private ArchitectureRuntimeDependencyFixture() {}

  /** Introduces the forbidden dependency for the architecture negative control. */
  public static Class<?> runtimeType() {
    return GridGrindRequestDoctor.class;
  }
}
