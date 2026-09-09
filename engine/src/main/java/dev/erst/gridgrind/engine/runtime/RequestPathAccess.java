package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/**
 * Request-scoped capability for every request-owned filesystem path.
 *
 * <p>Read inputs are copied through no-follow bindings into executor-private temporary files during
 * phase-four preflight. A persistence target retains its verified parent binding until the staged
 * artifact is committed. The original request path is never reopened after preflight.
 */
final class RequestPathAccess implements AutoCloseable {
  private final Path executionRoot;
  private final TempFileFactory tempFileFactory;
  private final GridGrindExecutionGrant executionGrant;

  // One request owns this capability on one execution thread; insertion order is
  // response-deterministic.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private final Map<String, MaterializedRead> materializedReads = new LinkedHashMap<>();

  private final List<Path> ownedMaterializations = new ArrayList<>();

  // One request owns this capability on one execution thread; insertion order is
  // response-deterministic.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private final Map<PathWarningKey, RequestWarning> absolutePathWarnings = new LinkedHashMap<>();

  private Optional<RequestPathBinding> outputBinding = Optional.empty();
  private Optional<String> outputPath = Optional.empty();

  RequestPathAccess(
      Path executionRoot, TempFileFactory tempFileFactory, GridGrindExecutionGrant executionGrant) {
    this.executionRoot =
        Objects.requireNonNull(executionRoot, "executionRoot must not be null")
            .toAbsolutePath()
            .normalize();
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    this.executionGrant = Objects.requireNonNull(executionGrant, "executionGrant must not be null");
  }

  /** Returns a private materialization of one verified request-owned input file. */
  Path materializeRead(String rawPath, String pathRole, String prefix, String suffix)
      throws IOException {
    Objects.requireNonNull(rawPath, "rawPath must not be null");
    Objects.requireNonNull(pathRole, "pathRole must not be null");
    Objects.requireNonNull(prefix, "prefix must not be null");
    Objects.requireNonNull(suffix, "suffix must not be null");
    MaterializedRead existing = materializedReads.get(rawPath);
    if (existing != null) {
      return existing.materializedPath();
    }

    try (RequestPathBinding binding = RequestPathBinding.bindExistingRead(rawPath, executionRoot)) {
      if (Files.isDirectory(binding.resolvedPath(), java.nio.file.LinkOption.NOFOLLOW_LINKS)
          && "source".equals(pathRole)) {
        throw new SourcePathIsDirectoryException(rawPath);
      }
      requireReadAuthority(binding.resolvedPath(), pathRole);
      recordAbsolutePathWarning(rawPath, pathRole, binding.resolvedPath());
      Path materialized = tempFileFactory.createTempFile(prefix, suffix);
      boolean completed = false;
      try {
        try (InputStream input = binding.openInputStream();
            OutputStream output = Files.newOutputStream(materialized)) {
          input.transferTo(output);
        }
        materializedReads.put(
            rawPath,
            new MaterializedRead(
                pathRole,
                binding.resolvedPath(),
                materialized,
                Files.size(materialized),
                RequestPathDigest.sha256(materialized)));
        ownedMaterializations.add(materialized);
        completed = true;
        return materialized;
      } finally {
        if (!completed) {
          deleteQuietly(materialized);
        }
      }
    }
  }

  /**
   * Binds the sole persistence target during phase-four preflight without creating or writing it.
   */
  @SuppressWarnings(
      "PMD.CloseResource") // Ownership transfers to outputBinding until this capability closes.
  void prepareOutput(String rawPath, String pathRole, WorkbookArtifactWriteDisposition disposition)
      throws IOException {
    Objects.requireNonNull(rawPath, "rawPath must not be null");
    Objects.requireNonNull(pathRole, "pathRole must not be null");
    Objects.requireNonNull(disposition, "disposition must not be null");
    if (outputBinding.isPresent()) {
      if (!outputPath.orElseThrow().equals(rawPath)) {
        throw new IllegalStateException("one request may prepare only one persistence target");
      }
      return;
    }
    RequestPathBinding binding = RequestPathBinding.bindWriteTarget(rawPath, executionRoot);
    recordAbsolutePathWarning(rawPath, pathRole, binding.resolvedPath());
    if (binding.hasExistingLeaf()
        && Files.isDirectory(binding.resolvedPath(), java.nio.file.LinkOption.NOFOLLOW_LINKS)) {
      binding.close();
      throw new OutputPathIsDirectoryException(rawPath);
    }
    if (disposition == WorkbookArtifactWriteDisposition.CREATE_NEW && binding.hasExistingLeaf()) {
      binding.close();
      throw new OutputPathAlreadyExistsException(rawPath);
    }
    outputBinding = Optional.of(binding);
    outputPath = Optional.of(rawPath);
  }

  /** Publishes one verified staged workbook through the phase-four-bound persistence parent. */
  PublicationOutcome.Published publishOutput(
      Path stagedFile,
      WorkbookArtifactWriteDisposition disposition,
      PublicationOutcome.StagedArtifactVerification stagedArtifactVerification)
      throws IOException {
    Objects.requireNonNull(stagedFile, "stagedFile must not be null");
    Objects.requireNonNull(disposition, "disposition must not be null");
    Objects.requireNonNull(
        stagedArtifactVerification, "stagedArtifactVerification must not be null");
    PublicationAttempt attempt =
        new RequestPathPublication()
            .publish(
                outputBinding.orElseThrow(
                    () ->
                        new IllegalStateException(
                            "a persistence write requires a phase-four-bound output target")),
                stagedFile,
                disposition,
                stagedArtifactVerification);
    return switch (attempt) {
      case PublishedPublicationAttempt published -> published.outcome();
      case FailedPublicationAttempt failed ->
          throw new WorkbookPublicationException(failed.outcome(), failed.failure());
    };
  }

  /**
   * Returns the normalized execution-root path used only for report context and internal staging.
   */
  Path executionRoot() {
    return executionRoot;
  }

  /** Returns the prepared report path for the persistence target. */
  Path outputPath() {
    return outputBinding
        .orElseThrow(() -> new IllegalStateException("no persistence target was prepared"))
        .resolvedPath();
  }

  List<RequestWarning> warnings() {
    return List.copyOf(absolutePathWarnings.values());
  }

  /** Returns immutable identities for every private input copy retained by this request. */
  List<MaterializedReadIdentity> materializedReadIdentities() {
    return materializedReads.values().stream()
        .map(
            read ->
                new MaterializedReadIdentity(
                    read.role(), read.resolvedPath().toString(), read.byteSize(), read.sha256()))
        .toList();
  }

  /** Returns the private immutable materialization previously bound for this exact request path. */
  Path materializedReadPath(String rawPath) {
    Objects.requireNonNull(rawPath, "rawPath must not be null");
    MaterializedRead materializedRead = materializedReads.get(rawPath);
    if (materializedRead == null) {
      throw new IllegalStateException(
          "request path was not materialized during admission: " + rawPath);
    }
    return materializedRead.materializedPath();
  }

  @Override
  public void close() throws IOException {
    List<Cleanup> cleanups = new ArrayList<>(ownedMaterializations.size() + 1);
    outputBinding.ifPresent(binding -> cleanups.add(binding::close));
    for (int index = ownedMaterializations.size() - 1; index >= 0; index--) {
      Path materialization = ownedMaterializations.get(index);
      cleanups.add(() -> Files.deleteIfExists(materialization));
    }
    closeAll(cleanups);
  }

  static void closeAll(List<Cleanup> cleanups) throws IOException {
    IOException failure = null;
    for (Cleanup cleanup : cleanups) {
      try {
        cleanup.close();
      } catch (IOException exception) {
        if (failure == null) {
          failure = exception;
        } else {
          failure.addSuppressed(exception);
        }
      }
    }
    if (failure != null) {
      throw failure;
    }
  }

  private static void deleteQuietly(Path path) {
    try {
      Files.deleteIfExists(path);
    } catch (IOException ignored) {
      // A failed preflight must not mask its primary diagnostic with private-temp cleanup noise.
    }
  }

  private void recordAbsolutePathWarning(String rawPath, String pathRole, Path normalizedPath) {
    if (!Path.of(rawPath).isAbsolute()) {
      return;
    }
    PathWarningKey key = new PathWarningKey(normalizedPath, pathRole);
    absolutePathWarnings.putIfAbsent(
        key, RequestWarning.nonPortableAbsolutePath(normalizedPath.toString(), pathRole));
  }

  private void requireReadAuthority(Path resolvedPath, String pathRole) {
    GridGrindExecutionGrant.Bounded bounded = (GridGrindExecutionGrant.Bounded) executionGrant;
    boolean allowed =
        bounded.readableResources().stream()
            .filter(GridGrindExecutionGrant.ReadAuthority.File.class::isInstance)
            .map(GridGrindExecutionGrant.ReadAuthority.File.class::cast)
            .anyMatch(resource -> resource.path().equals(resolvedPath));
    if (!allowed) {
      throw new ExecutionAuthorityDeniedException(
          "host grant does not permit reading " + pathRole + ": " + resolvedPath);
    }
  }

  private record PathWarningKey(Path path, String pathRole) {
    private PathWarningKey {
      Objects.requireNonNull(path, "path must not be null");
      Objects.requireNonNull(pathRole, "pathRole must not be null");
    }
  }

  private record MaterializedRead(
      String role, Path resolvedPath, Path materializedPath, long byteSize, String sha256) {}

  record MaterializedReadIdentity(String role, String path, long byteSize, String sha256) {}

  /**
   * Closes or deletes one request-private resource while retaining checked I/O failure semantics.
   */
  @FunctionalInterface
  interface Cleanup {
    /** Releases this resource. */
    void close() throws IOException;
  }
}
