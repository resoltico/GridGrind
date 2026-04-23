package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Focused coverage for request-rooted execution binding helpers. */
class CliExecutionBindingsFactoryTest {
  @Test
  void executionWorkingDirectoryPreservesRootPathsThatHaveNoParent() {
    Path root = Path.of("").toAbsolutePath().normalize().getRoot();
    assertNotNull(root);

    assertEquals(root, CliExecutionBindingsFactory.executionWorkingDirectory(root));
  }
}
