package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Owns one run's resources, snapshots and accumulated facts until finalization. */
final class ExecutionSession implements AutoCloseable {
  private final java.util.List<RequestPathAccess.Cleanup> cleanups = new java.util.ArrayList<>();
  private final Path parent;
  private @Nullable Path root;

  final java.util.List<dev.erst.gridgrind.core.problem.RequestWarning> warnings =
      new java.util.ArrayList<>();
  final java.util.List<dev.erst.gridgrind.core.plan.AssertionResult> planAssertions =
      new java.util.ArrayList<>();
  final java.util.List<dev.erst.gridgrind.core.plan.AssertionResult> hostAssertions =
      new java.util.ArrayList<>();
  final java.util.List<dev.erst.gridgrind.core.plan.InspectionResult> inspections =
      new java.util.ArrayList<>();
  final java.util.List<dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation>
      preservation = new java.util.ArrayList<>();
  dev.erst.gridgrind.core.grant.AcceptancePolicy acceptancePolicy;
  java.util.Optional<RequestPreflight.PreparedRequest> prepared = java.util.Optional.empty();
  java.util.Optional<dev.erst.gridgrind.core.model.formula.CalculationReport> calculation =
      java.util.Optional.empty();
  final InputSnapshots inputs;
  final dev.erst.gridgrind.engine.media.MediaAdmissionBudget media =
      new dev.erst.gridgrind.engine.media.MediaAdmissionBudget();
  final PublicationFacts publication = new PublicationFacts();

  ExecutionSession(Path parent, dev.erst.gridgrind.core.grant.ExecutionGrant grant) {
    this.inputs = new InputSnapshots(grant);
    this.acceptancePolicy = grant.acceptancePolicy();
    this.parent = Objects.requireNonNull(parent).toAbsolutePath().normalize();
  }

  private java.util.Optional<
          dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.ExecutionStrategy>
      strategy = java.util.Optional.empty();

  void recordExecutionStrategy(
      dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.ExecutionStrategy value) {
    if (strategy.isPresent())
      throw new IllegalStateException("Execution strategy is already selected");
    strategy = java.util.Optional.of(Objects.requireNonNull(value));
  }

  java.util.Optional<dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.ExecutionStrategy>
      executionStrategy() {
    return strategy;
  }

  void prepared(RequestPreflight.PreparedRequest value) {
    if (prepared.isPresent()) throw new IllegalStateException("Execution is already admitted");
    prepared = java.util.Optional.of(Objects.requireNonNull(value));
  }

  Path createTempFile(String prefix, String suffix) throws IOException {
    if (root == null) {
      Files.createDirectories(parent);
      root = Files.createTempDirectory(parent, "gridgrind-run-");
    }
    return Files.createTempFile(root, prefix, suffix);
  }

  void own(RequestPathAccess.Cleanup cleanup) {
    cleanups.add(Objects.requireNonNull(cleanup));
  }

  @Override
  public void close() throws IOException {
    IOException failure = publication.cleanupFailure();
    for (int index = cleanups.size() - 1; index >= 0; index--) {
      try {
        cleanups.get(index).close();
      } catch (IOException exception) {
        failure = appendCleanupFailure(failure, exception);
      } catch (RuntimeException exception) {
        failure = appendCleanupFailure(failure, cleanupFailure(exception));
      }
    }
    cleanups.clear();
    if (root != null) {
      try {
        dev.erst.gridgrind.engine.packaging.WorkbookFiles.deleteTreeIfExists(root);
      } catch (IOException exception) {
        failure = appendCleanupFailure(failure, exception);
      } catch (RuntimeException exception) {
        failure = appendCleanupFailure(failure, cleanupFailure(exception));
      }
    }
    if (failure != null) throw failure;
    root = null;
  }

  private static IOException cleanupFailure(RuntimeException failure) {
    return new IOException("Run cleanup failed", failure);
  }

  @SuppressWarnings({"ReferenceEquality", "PMD.CompareObjectsWithEquals"})
  // Throwable self-suppression is forbidden by identity, even for value-equal subclasses.
  private static IOException appendCleanupFailure(
      @Nullable IOException primary, IOException secondary) {
    if (primary == null) return secondary;
    if (primary != secondary && primary.getSuppressed().length < 8)
      primary.addSuppressed(secondary);
    return primary;
  }
}
