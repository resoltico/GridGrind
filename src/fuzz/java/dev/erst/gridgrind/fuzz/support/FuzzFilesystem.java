package dev.erst.gridgrind.fuzz.support;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.util.Comparator;

/** Deletes fuzz-owned scratch; cleanup failures must not silently accumulate files. */
public final class FuzzFilesystem {
  private FuzzFilesystem() {}

  /** Deletes one owned tree without following symbolic links. */
  public static void deleteRecursively(Path root) {
    if (root == null || !Files.exists(root, LinkOption.NOFOLLOW_LINKS)) return;
    try (var paths = Files.walk(root)) {
      for (Path path : paths.sorted(Comparator.reverseOrder()).toList()) Files.delete(path);
    } catch (IOException failure) {
      throw new UncheckedIOException("Cannot clean fuzz-owned scratch: " + root, failure);
    }
  }

  /** Retains the original execution or generation failure when scratch cleanup also fails. */
  public static void cleanupAfterFailure(Path root, Throwable failure) {
    try {
      deleteRecursively(root);
    } catch (RuntimeException cleanupFailure) {
      failure.addSuppressed(cleanupFailure);
    }
  }
}
