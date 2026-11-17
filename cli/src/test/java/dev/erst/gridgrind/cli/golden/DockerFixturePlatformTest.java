package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Rejects disagreement between requested, inspected and independently observed JVM platforms. */
class DockerFixturePlatformTest {
  @Test
  void onlyTheTwoSupportedReleasePlatformsCanBeRequested() {
    assertEquals("", DockerFixture.validateRequestedPlatform(""));
    assertEquals("linux/amd64", DockerFixture.validateRequestedPlatform("linux/amd64"));
    assertEquals("linux/arm64", DockerFixture.validateRequestedPlatform("linux/arm64"));
    assertThrows(
        IllegalArgumentException.class,
        () -> DockerFixture.validateRequestedPlatform("linux/arm/v7"));
    assertThrows(
        IllegalArgumentException.class,
        () -> DockerFixture.validateRequestedPlatform(" linux/amd64"));
  }

  @Test
  void aHostArchitectureInspectionCannotSatisfyTheOtherRequestedArchitecture() {
    assertEquals("linux/arm64", DockerFixture.requirePlatform("linux/arm64", "linux/arm64"));
    assertEquals("linux/amd64", DockerFixture.requirePlatform("", "linux/amd64"));
    assertThrows(
        IllegalStateException.class,
        () -> DockerFixture.requirePlatform("linux/arm64", "linux/amd64"));
    assertThrows(
        IllegalStateException.class,
        () -> DockerFixture.requirePlatform("linux/amd64", "linux/arm64"));
  }

  @Test
  void matchingImageMetadataCannotHideTheWrongExecutedJvm() {
    DockerFixture.requireJvmPlatform("linux/arm64", "    os.name = Linux\n    os.arch = aarch64\n");
    DockerFixture.requireJvmPlatform("linux/amd64", "    os.arch = amd64\n    os.name = Linux\n");
    assertThrows(
        IllegalStateException.class,
        () ->
            DockerFixture.requireJvmPlatform("linux/arm64", "os.arch = amd64\nos.name = Linux\n"));
    assertThrows(
        IllegalStateException.class,
        () ->
            DockerFixture.requireJvmPlatform(
                "linux/amd64", "os.arch = aarch64\nos.name = Linux\n"));
  }

  @Test
  void absentOrConflictingJvmFactsAreRefused() {
    assertThrows(
        IllegalStateException.class,
        () -> DockerFixture.requireJvmPlatform("linux/amd64", "os.name = Linux\n"));
    assertThrows(
        IllegalStateException.class,
        () ->
            DockerFixture.requireJvmPlatform(
                "linux/amd64", "os.arch = amd64\nos.name = Windows\n"));
    assertThrows(
        IllegalStateException.class,
        () ->
            DockerFixture.requireJvmPlatform(
                "linux/arm64", "os.arch = aarch64\nos.arch = amd64\nos.name = Linux\n"));
  }
}
