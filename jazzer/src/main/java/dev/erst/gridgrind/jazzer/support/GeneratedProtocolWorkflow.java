package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.protocol.GridGrindRequest;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;

/** Carries one generated protocol workflow plus any local scratch paths it created. */
public record GeneratedProtocolWorkflow(GridGrindRequest request, List<Path> cleanupRoots) {
  public GeneratedProtocolWorkflow {
    Objects.requireNonNull(request, "request must not be null");
    cleanupRoots = cleanupRoots == null ? List.of() : List.copyOf(cleanupRoots);
    for (Path cleanupRoot : cleanupRoots) {
      Objects.requireNonNull(cleanupRoot, "cleanupRoots must not contain nulls");
    }
  }

  /** Deletes every generated local scratch directory or file owned by this workflow. */
  public void cleanup() {
    cleanupRoots.forEach(GeneratedProtocolWorkflow::deleteRecursively);
  }

  private static void deleteRecursively(Path root) {
    try {
      if (!Files.exists(root)) {
        return;
      }
      Files.walk(root)
          .sorted(Comparator.reverseOrder())
          .forEach(
              path -> {
                try {
                  Files.deleteIfExists(path);
                } catch (IOException ignored) {
                  // Best-effort scratch cleanup only.
                }
              });
    } catch (IOException ignored) {
      // Best-effort scratch cleanup only.
    }
  }
}
