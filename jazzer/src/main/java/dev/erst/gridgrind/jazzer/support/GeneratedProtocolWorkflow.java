package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Carries one generated protocol workflow plus any local scratch paths it created. */
public record GeneratedProtocolWorkflow(WorkbookPlan request, List<Path> cleanupRoots) {
  public GeneratedProtocolWorkflow {
    Objects.requireNonNull(request, "request must not be null");
    cleanupRoots = cleanupRoots == null ? List.of() : List.copyOf(cleanupRoots);
    for (Path cleanupRoot : cleanupRoots) {
      Objects.requireNonNull(cleanupRoot, "cleanupRoots must not contain nulls");
    }
  }

  /** Deletes every generated local scratch directory or file owned by this workflow. */
  public void cleanup() {
    cleanupRoots.forEach(JazzerFilesystemSupport::deleteRecursively);
  }

  /** Returns the explicit execution root that owns this generated workflow's relative paths. */
  public Path executionRoot() {
    if (cleanupRoots.isEmpty()) {
      throw new IllegalStateException("generated workflow must declare one execution root");
    }
    return cleanupRoots.getFirst().toAbsolutePath().normalize();
  }
}
