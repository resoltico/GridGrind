package dev.erst.gridgrind.engine.packaging;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies complete owned-tree cleanup attempts, bounded failures and no-follow containment. */
class WorkbookFilesCleanupTest {
  @TempDir Path root;

  @Test
  void deletionContinuesAfterFailuresAndRetainsAtMostEightSecondaryCauses() throws Exception {
    for (int index = 0; index < 12; index++)
      Files.writeString(root.resolve("file-" + index), "owned");
    var attempted = new HashSet<Path>();
    var first = new IOException("synthetic deletion failure");
    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                WorkbookFiles.deleteTreeIfExists(
                    root,
                    path -> {
                      attempted.add(path);
                      if (attempted.size() == 1) throw first;
                      throw new IOException("another deletion failure");
                    }));
    assertSame(first, failure);
    assertEquals(13, attempted.size());
    assertEquals(8, failure.getSuppressed().length);
    WorkbookFiles.deleteTreeIfExists(root);
    assertFalse(Files.exists(root));
  }

  @Test
  void treeCleanupDoesNotFollowOwnedSymlinksIntoHostFiles() throws Exception {
    Path host = Files.createDirectory(root.resolve("host"));
    Files.writeString(host.resolve("keep"), "host-owned");
    Path owned = Files.createDirectory(root.resolve("owned"));
    Files.createSymbolicLink(owned.resolve("link"), host);
    WorkbookFiles.deleteTreeIfExists(owned);
    assertFalse(Files.exists(owned));
    assertTrue(Files.exists(host.resolve("keep")));
    WorkbookFiles.deleteTreeIfExists(root.resolve("missing"));
  }
}
