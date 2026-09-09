package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;

/** Host-required checks applied only after the staged artifact has reopened successfully. */
@FunctionalInterface
interface StagedArtifactAcceptance {
  /** Verifies one decrypted, structurally reopened staged workbook before publication. */
  void verify(Path materializedStagedWorkbook) throws IOException, PreservationFailedException;

  /** Returns the explicit no-additional-requirements policy for a workflow with no such grant. */
  static StagedArtifactAcceptance none() {
    return materializedStagedWorkbook ->
        Objects.requireNonNull(
            materializedStagedWorkbook, "materializedStagedWorkbook must not be null");
  }
}
