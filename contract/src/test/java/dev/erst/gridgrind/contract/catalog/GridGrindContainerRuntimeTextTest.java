package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Focused coverage for canonical container-runtime wording surfaces. */
class GridGrindContainerRuntimeTextTest {
  @Test
  void mountedWorkdirSurfacesStayCanonical() {
    assertEquals("/work", GridGrindContainerRuntimeText.containerWorkdirPath());
    assertEquals(
        "--user \"$(id -u):$(id -g)\"",
        GridGrindContainerRuntimeText.dockerMountedWorkdirUserArgument());
    assertEquals(
        "-v \"$(pwd)\":/work", GridGrindContainerRuntimeText.dockerMountedWorkdirVolumeArgument());
    assertEquals(
        "docker run --rm -i --user \"$(id -u):$(id -g)\" -v \"$(pwd)\":/work"
            + " ghcr.io/resoltico/gridgrind:latest --request request.json --response"
            + " response.json",
        GridGrindContainerRuntimeText.dockerMountedWorkdirExecutionCommand(
            "ghcr.io/resoltico/gridgrind:latest"));
    assertTrue(
        GridGrindContainerRuntimeText.dockerMountedWorkdirSummary()
            .contains(GridGrindContainerRuntimeText.containerWorkdirPath()));
    assertTrue(
        GridGrindContainerRuntimeText.dockerMountedWorkdirSummary()
            .contains(GridGrindContainerRuntimeText.dockerMountedWorkdirUserArgument()));
    assertTrue(
        GridGrindContainerRuntimeText.dockerMountedWorkdirSummary().contains("rootless runtime"));
  }
}
