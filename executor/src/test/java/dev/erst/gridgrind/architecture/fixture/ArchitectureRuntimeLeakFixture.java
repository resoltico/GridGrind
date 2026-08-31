package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor;
import java.util.List;
import java.util.Objects;

/** Deliberately invalid public signatures used to prove implementation leaks are rejected. */
public final class ArchitectureRuntimeLeakFixture {
  private ArchitectureRuntimeLeakFixture() {}

  /** Echoes a forbidden runtime type solely for architecture-rule regression. */
  public static GridGrindRequestDoctor leakedRuntimeType(GridGrindRequestDoctor runtimeType) {
    return Objects.requireNonNull(runtimeType, "runtimeType must not be null");
  }

  /** Preserves a generic runtime type solely for architecture-rule regression. */
  public static List<GridGrindRequestDoctor> leakedGenericRuntimeType(
      List<GridGrindRequestDoctor> runtimeTypes) {
    return List.copyOf(runtimeTypes);
  }
}
