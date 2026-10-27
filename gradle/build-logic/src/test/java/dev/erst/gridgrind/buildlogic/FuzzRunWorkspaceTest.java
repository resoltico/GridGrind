package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises actual corpus lifetime, seed admission, pending findings and exclusive ownership. */
class FuzzRunWorkspaceTest {
  @TempDir Path root;

  @Test
  void retainsEverySeedAndOnlyTheLatestCorpus() throws Exception {
    Path seeds = Files.createDirectories(root.resolve("seeds"));
    Files.write(seeds.resolve("one"), new byte[] {1, 2});
    Path workspace = root.resolve("work");
    AtomicReference<Path> first = new AtomicReference<>();
    FuzzRunWorkspace.run(workspace, seeds, "protocol-request", 100, 32, files -> {
      first.set(files.getCorpus());
      assertTrue(Files.isRegularFile(files.getCorpus().resolve("one")));
      assertEquals(3202L, files.getCorpusByteCeiling());
      assertTrue(Files.isDirectory(files.getTemporary()));
    });
    assertFalse(Files.exists(workspace.resolve("tmp")));
    FuzzRunWorkspace.run(workspace, seeds, "xlsx-roundtrip", 100, 32, files ->
        assertTrue(Files.isRegularFile(files.getCorpus().resolve("one"))));
    assertFalse(Files.exists(first.get()));
  }

  @Test
  void refusingAnOversizedSeedDoesNotEraseThePriorCorpus() throws Exception {
    Path seeds = Files.createDirectories(root.resolve("seeds"));
    Files.write(seeds.resolve("one"), new byte[] {1, 2});
    Path workspace = root.resolve("work");
    Path retained = Files.createDirectories(workspace.resolve("corpus")).resolve("retained");
    Files.writeString(retained, "keep");
    assertThrows(IllegalArgumentException.class,
        () -> FuzzRunWorkspace.run(workspace, seeds, "protocol-request", 100, 1, files -> {}));
    assertEquals("keep", Files.readString(retained));
  }

  @Test
  void pendingFindingsBlockCleanupAndFurtherRuns() throws Exception {
    Path seeds = Files.createDirectories(root.resolve("seeds"));
    Files.writeString(seeds.resolve("one"), "seed");
    Path workspace = root.resolve("work");
    Path finding = Files.createDirectories(workspace.resolve("findings/protocol-request")).resolve("crash");
    Files.writeString(finding, "keep");
    assertThrows(IllegalStateException.class,
        () -> FuzzRunWorkspace.run(workspace, seeds, "xlsx-roundtrip", 100, 32, files -> {}));
    assertEquals("keep", Files.readString(finding));
  }

  @Test
  void concurrentRunCannotResetTheActiveCorpus() throws Exception {
    Path seeds = Files.createDirectories(root.resolve("seeds"));
    Files.writeString(seeds.resolve("one"), "seed");
    Path workspace = root.resolve("work");
    FuzzRunWorkspace.run(workspace, seeds, "protocol-request", 100, 32, outer -> {
      assertThrows(IllegalStateException.class,
          () -> FuzzRunWorkspace.run(workspace, seeds, "xlsx-roundtrip", 100, 32, inner -> {}));
      assertTrue(Files.exists(outer.getCorpus().resolve("one")));
    });
  }
}
