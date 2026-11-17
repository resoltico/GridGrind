package dev.erst.gridgrind.cli.io;

import dev.erst.gridgrind.cli.command.CliPathArguments;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Builds one CLI execution binding plus one owned private scratch directory. */
public final class CliExecutionBindingsFactory {
  private static final String SYSTEM_TEMP_PROPERTY = "java.io.tmpdir";
  private static final String MANAGED_TEMP_ROOT_PREFIX = "gridgrind-cli-";

  private CliExecutionBindingsFactory() {}

  /** Binds explicit authority, request paths, secrets and one owned scratch directory. */
  public static ManagedRequestInputs create(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      Optional<Path> grantPath,
      Optional<Path> secretsProviderPath,
      WorkbookPlan request,
      InputStream stdin)
      throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Objects.requireNonNull(grantPath, "grantPath must not be null");
    Objects.requireNonNull(secretsProviderPath, "secretsProviderPath must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Path workingDirectory = executionWorkingDirectory(requestPath, executionRootPath);
    ExecutionGrant executionGrant =
        grantPath.isPresent() ? readAuthority(grantPath.orElseThrow()) : ExecutionGrant.denied();
    Optional<CliFileSecretResolver> secretResolver =
        secretsProviderPath.map(path -> {
          try {
            return CliFileSecretResolver.open(path);
          } catch (IOException exception) {
            throw new CliSecretReadException(path.toString(), exception);
          }
        });
    ManagedTempRoot tempRoot;
    try {
      tempRoot = createManagedTempRoot(tempRootPath);
    } catch (IOException | RuntimeException | Error failure) {
      var failures = new CleanupFailures(failure);
      failures.run(() -> {
        if (secretResolver.isPresent()) secretResolver.orElseThrow().close();
      });
      failures.rethrow();
      throw failure;
    }
    try {
      GridGrindRequestInputs inputs;
      if (GridGrindRequestRequirements.requiresStandardInput(request)
          || GridGrindRequestRequirements.requiresStandardInput(executionGrant)) {
        dev.erst.gridgrind.engine.api.GridGrindPayloadOpener payload =
            () ->
                new java.io.FilterInputStream(stdin) {
                  @Override
                  public void close() {
                    // System.in and other process-owned CLI streams are borrowed.
                  }
                };
        inputs =
            new GridGrindRequestInputs(
                workingDirectory,
                tempRoot.root(),
                payload,
                executionGrant,
                secretResolver.map(
                    dev.erst.gridgrind.engine.api.GridGrindSecretResolver.class::cast));
      } else {
        inputs =
            secretResolver
                .<GridGrindRequestInputs>map(resolver ->
                    new GridGrindRequestInputs(
                        workingDirectory, tempRoot.root(), executionGrant, resolver))
                .orElseGet(() ->
                    new GridGrindRequestInputs(workingDirectory, tempRoot.root(), executionGrant));
      }
      return new ManagedRequestInputs(inputs, secretResolver);
    } catch (RuntimeException | Error failure) {
      var failures = new CleanupFailures(failure);
      failures.run(() -> {
        if (secretResolver.isPresent()) secretResolver.orElseThrow().close();
      });
      failures.run(() -> deleteTreeIfExists(tempRoot.root()));
      failures.rethrow();
      throw failure;
    }
  }

  private static ExecutionGrant readAuthority(Path supplied) throws IOException {
    Path path = supplied.toAbsolutePath().normalize();
    try {
      return dev.erst.gridgrind.json.codec.GridGrindJson.readGrant(
          Files.readAllBytes(path), Path.of(""));
    } catch (dev.erst.gridgrind.core.problem.InvalidAuthorityException refusal) {
      throw new CliAuthorityDocumentException(path, refusal);
    } catch (IOException io) {
      throw new CliAuthorityIOException(path, io);
    } catch (RuntimeException unexpected) {
      throw new CliAuthorityUnexpectedException(path, unexpected);
    }
  }

  static Path executionWorkingDirectory(
      Optional<Path> requestPath, Optional<Path> executionRootPath) {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    if (requestPath.isEmpty() || CliPathArguments.isStandardInputPath(requestPath)) {
      return executionRootPath
          .orElseThrow(() -> new IllegalArgumentException("executionRootPath must be present"))
          .toAbsolutePath()
          .normalize();
    }
    Path normalizedRequestPath = requestPath.orElseThrow().toAbsolutePath().normalize();
    Path parent = normalizedRequestPath.getParent();
    return parent == null ? normalizedRequestPath : parent;
  }

  static Path tempRootParent(Optional<Path> tempRootPath) throws IOException {
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    if (tempRootPath.isPresent()) {
      return tempRootPath.orElseThrow().toAbsolutePath().normalize();
    }
    return defaultTempRootParent();
  }

  static ManagedTempRoot createManagedTempRoot(Optional<Path> tempRootPath) throws IOException {
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Path tempRootParent = tempRootParent(tempRootPath);
    Files.createDirectories(tempRootParent);
    Path tempRoot =
        Files.createTempDirectory(tempRootParent, MANAGED_TEMP_ROOT_PREFIX)
            .toAbsolutePath()
            .normalize();
    return new ManagedTempRoot(tempRoot);
  }

  static void deleteTreeIfExists(Path root) throws IOException {
    if (root == null) return;
    var failures = new CleanupFailures(null);
    var visitor =
        new java.nio.file.SimpleFileVisitor<Path>() {
          @Override
          public java.nio.file.FileVisitResult visitFile(
              Path file, java.nio.file.attribute.BasicFileAttributes attributes) {
            failures.run(() -> Files.deleteIfExists(file));
            return java.nio.file.FileVisitResult.CONTINUE;
          }

          @Override
          public java.nio.file.FileVisitResult visitFileFailed(Path file, IOException exception) {
            if (!(exception instanceof java.nio.file.NoSuchFileException))
              failures.record(exception);
            return java.nio.file.FileVisitResult.CONTINUE;
          }

          @Override
          public java.nio.file.FileVisitResult postVisitDirectory(
              Path directory, @org.jspecify.annotations.Nullable IOException exception) {
            if (exception != null) failures.record(exception);
            failures.run(() -> Files.deleteIfExists(directory));
            return java.nio.file.FileVisitResult.CONTINUE;
          }
        };
    failures.run(() -> Files.walkFileTree(root, visitor));
    failures.rethrow();
  }

  /** Retains cleanup diagnostics while every owned resource unwinds and keeps Errors fatal. */
  private static final class CleanupFailures {
    private @org.jspecify.annotations.Nullable Throwable failure;

    private CleanupFailures(@org.jspecify.annotations.Nullable Throwable failure) {
      this.failure = failure;
    }

    private void run(Cleanup cleanup) {
      try {
        cleanup.close();
      } catch (IOException | RuntimeException | Error exception) {
        record(exception);
      }
    }

    @SuppressWarnings({"ReferenceEquality", "PMD.CompareObjectsWithEquals"})
    // Throwable self-suppression is forbidden by identity, independent of equals implementations.
    private void record(Throwable exception) {
      if (failure == null) {
        failure = exception;
        return;
      }
      if (failure == exception) return;
      Throwable primary = failure;
      Throwable secondary = exception;
      if (exception instanceof Error && !(primary instanceof Error)) {
        secondary = primary;
        primary = exception;
        failure = primary;
      }
      if (primary.getSuppressed().length < 8) primary.addSuppressed(secondary);
    }

    private void rethrow() throws IOException {
      if (failure instanceof IOException io) throw io;
      if (failure instanceof RuntimeException runtime) throw runtime;
      if (failure instanceof Error error) throw error;
      if (failure != null) throw new AssertionError("Unexpected cleanup failure", failure);
    }
  }

  /** One cleanup obligation for a CLI-owned descriptor or scratch tree. */
  @FunctionalInterface
  private interface Cleanup {
    /** Releases the owned resource, retaining checked filesystem failure semantics. */
    void close() throws IOException;
  }

  private static Path defaultTempRootParent() throws IOException {
    String systemTempDir = System.getProperty(SYSTEM_TEMP_PROPERTY);
    if (systemTempDir == null || systemTempDir.isBlank()) {
      throw new IOException("System temporary-file root is unavailable");
    }
    return Path.of(systemTempDir).toAbsolutePath().normalize();
  }

  /** One CLI-owned execution binding plus enough metadata to clean its scratch root on close. */
  public static final class ManagedRequestInputs implements AutoCloseable {
    private final GridGrindRequestInputs inputs;
    private final Optional<CliFileSecretResolver> secretResolver;

    ManagedRequestInputs(
        GridGrindRequestInputs inputs, Optional<CliFileSecretResolver> secretResolver) {
      this.inputs = Objects.requireNonNull(inputs, "inputs must not be null");
      this.secretResolver =
          Objects.requireNonNull(secretResolver, "secretResolver must not be null");
    }

    /** Returns the host inputs retained for this invocation. */
    public GridGrindRequestInputs inputs() {
      return inputs;
    }

    @Override
    public void close() throws IOException {
      var failures = new CleanupFailures(null);
      failures.run(() -> {
        if (secretResolver.isPresent()) secretResolver.orElseThrow().close();
      });
      failures.run(() -> deleteTreeIfExists(inputs.tempRoot()));
      failures.rethrow();
    }
  }

  public record ManagedTempRoot(Path root) {
    public ManagedTempRoot {
      Objects.requireNonNull(root, "root must not be null");
    }
  }
}
