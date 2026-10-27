package dev.erst.gridgrind.cli.render;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Focused coverage for canonical container-runtime wording surfaces. */
class ContainerHelpTest {
  @Test
  void mountedWorkdirSurfacesStayCanonical() {
    assertEquals("/work", ContainerHelp.containerWorkdirPath());
    assertEquals("--user \"$(id -u):$(id -g)\"", ContainerHelp.dockerMountedWorkdirUserArgument());
    assertEquals("-v \"$(pwd)\":/work", ContainerHelp.dockerMountedWorkdirVolumeArgument());
    assertEquals(
        "docker run --rm -i --user \"$(id -u):$(id -g)\" -v \"$(pwd)\":/work"
            + " ghcr.io/resoltico/gridgrind:latest run --request request.json --grant grant.json --response"
            + " response.json",
        ContainerHelp.dockerMountedWorkdirExecutionCommand("ghcr.io/resoltico/gridgrind:latest"));
    assertTrue(
        ContainerHelp.dockerMountedWorkdirSummary().contains(ContainerHelp.containerWorkdirPath()));
    assertTrue(
        ContainerHelp.dockerMountedWorkdirSummary()
            .contains(ContainerHelp.dockerMountedWorkdirUserArgument()));
    assertTrue(ContainerHelp.dockerMountedWorkdirSummary().contains("rootless runtime"));
  }
}
