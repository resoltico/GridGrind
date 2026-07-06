package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Coverage for explicit execution bindings after ambient runtime defaults were removed. */
class ExecutionInputBindingsTest {
  @Test
  void bindingsNormalizeAndExposeTheirExplicitTempRoot() {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "..", "tmp", "request-root"), Path.of("tmp", "..", "tmp", "temp-root"));

    assertTrue(bindings.workingDirectory().isAbsolute());
    assertEquals(Path.of("tmp", "temp-root").toAbsolutePath().normalize(), bindings.tempRoot());
  }

  @Test
  void tempFileFactoryCreatesArtifactsUnderTheExplicitTempRoot() throws IOException {
    Path tempRoot = Files.createTempDirectory("gridgrind-bindings-temp-root-");
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(Files.createTempDirectory("gridgrind-bindings-root-"), tempRoot);

    Path created = bindings.tempFileFactory().createTempFile("binding-", ".tmp");

    assertEquals(tempRoot.toAbsolutePath().normalize(), created.getParent());
  }
}
