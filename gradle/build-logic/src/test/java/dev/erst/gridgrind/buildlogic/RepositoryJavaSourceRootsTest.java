package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

class RepositoryJavaSourceRootsTest {
  @Test
  void discoversRepoOwnedJavaRootsAndSkipsEphemeralTrees() throws IOException {
    Path repositoryRoot = Files.createTempDirectory("gridgrind-repository-roots");
    Files.createDirectories(repositoryRoot.resolve("engine/src/main/java"));
    Files.createDirectories(repositoryRoot.resolve("engine/src/test/java"));
    Files.createDirectories(repositoryRoot.resolve("engine/src/testFixtures/java"));
    Files.createDirectories(repositoryRoot.resolve("executor/src/parityTest/java"));
    Files.createDirectories(repositoryRoot.resolve("jazzer/src/fuzz/java"));
    Files.createDirectories(repositoryRoot.resolve("gradle/build-logic/src/test/java"));
    Files.createDirectories(repositoryRoot.resolve("tmp/release-bootstrap/untracked/cli/src/main/java"));
    Files.createDirectories(repositoryRoot.resolve("cli/build/generated/src/main/java"));
    Files.createDirectories(repositoryRoot.resolve(".gradle/sandbox/src/main/java"));

    List<Path> discoveredRoots = RepositoryJavaSourceRoots.discover(repositoryRoot);

    assertEquals(
        List.of(
            "engine/src/main/java",
            "engine/src/test/java",
            "engine/src/testFixtures/java",
            "executor/src/parityTest/java",
            "gradle/build-logic/src/test/java",
            "jazzer/src/fuzz/java"),
        RepositoryJavaSourceRoots.relativePaths(repositoryRoot, discoveredRoots));
  }
}
