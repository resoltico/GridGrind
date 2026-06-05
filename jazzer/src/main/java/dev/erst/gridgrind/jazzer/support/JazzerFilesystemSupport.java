package dev.erst.gridgrind.jazzer.support;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;

/** Best-effort filesystem cleanup for Jazzer-owned scratch directories and files. */
public final class JazzerFilesystemSupport {
  private JazzerFilesystemSupport() {}

  /** Deletes one Jazzer-owned path tree when it exists. */
  public static void deleteRecursively(Path root) {
    if (root == null) {
      return;
    }
    try {
      if (!Files.exists(root)) {
        return;
      }
      try (var stream = Files.walk(root)) {
        stream
            .sorted(Comparator.reverseOrder())
            .forEach(
                path -> {
                  try {
                    Files.deleteIfExists(path);
                  } catch (IOException ignored) {
                    // Best-effort scratch cleanup only.
                  }
                });
      }
    } catch (IOException ignored) {
      // Best-effort scratch cleanup only.
    }
  }
}
