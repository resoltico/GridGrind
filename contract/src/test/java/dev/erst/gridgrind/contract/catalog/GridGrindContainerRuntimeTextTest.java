package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Focused coverage for canonical container-runtime wording fragments. */
class GridGrindContainerRuntimeTextTest {
  @Test
  void mountedWorkdirFragmentsStayCanonical() {
    assertEquals("/work", GridGrindContainerRuntimeText.containerWorkdirPath());
    assertEquals(
        "-v \"$(pwd)\":/work", GridGrindContainerRuntimeText.dockerMountedWorkdirVolumeArgument());
    assertTrue(
        GridGrindContainerRuntimeText.dockerMountedWorkdirSummary()
            .contains(GridGrindContainerRuntimeText.containerWorkdirPath()));
  }
}
