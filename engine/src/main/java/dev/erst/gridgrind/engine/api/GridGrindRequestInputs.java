package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class GridGrindRequestInputs {
  private final Path workingDirectory;
  private final Path tempRoot;
  private final Optional<GridGrindPayloadOpener> standardInputSource;
  private final ExecutionGrant executionGrant;
  private final Optional<GridGrindSecretResolver> secretResolver;

  /** Creates execution inputs from one working directory, temp root, and explicit host grant. */
  public GridGrindRequestInputs(
      Path workingDirectory, Path tempRoot, ExecutionGrant executionGrant) {
    this(
        workingDirectory,
        tempRoot,
        Optional.empty(),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.empty());
  }

  /**
   * Creates execution inputs from one working directory, temp root, stdin payload, and host grant.
   */
  public GridGrindRequestInputs(
      Path workingDirectory,
      Path tempRoot,
      byte[] standardInputBytes,
      ExecutionGrant executionGrant) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(bytesSource(standardInputBytes, executionGrant)),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.empty());
  }

  /** Creates execution inputs with standard input, host authority, and a secret resolver. */
  public GridGrindRequestInputs(
      Path workingDirectory,
      Path tempRoot,
      byte[] standardInputBytes,
      ExecutionGrant executionGrant,
      GridGrindSecretResolver secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(bytesSource(standardInputBytes, executionGrant)),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.of(Objects.requireNonNull(secretResolver, "secretResolver")));
  }

  /** Creates execution inputs with a host-owned secret resolver. */
  public GridGrindRequestInputs(
      Path workingDirectory,
      Path tempRoot,
      ExecutionGrant executionGrant,
      GridGrindSecretResolver secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        Optional.empty(),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.of(Objects.requireNonNull(secretResolver, "secretResolver")));
  }

  private GridGrindRequestInputs(
      Path workingDirectory,
      Path tempRoot,
      Optional<GridGrindPayloadOpener> standardInputSource,
      ExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(standardInputSource, "standardInputSource must not be null");
    Objects.requireNonNull(executionGrant, "executionGrant must not be null");
    Objects.requireNonNull(secretResolver, "secretResolver must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.tempRoot = tempRoot.toAbsolutePath().normalize();
    this.standardInputSource = standardInputSource;
    this.executionGrant = executionGrant;
    this.secretResolver = secretResolver;
  }

  /** Returns the normalized working directory used to resolve relative authored input paths. */
  public Path workingDirectory() {
    return workingDirectory;
  }

  /** Returns the normalized temp root used for one execution's internal scratch files. */
  public Path tempRoot() {
    return tempRoot;
  }

  /** Returns true when stdin bytes are available to STANDARD_INPUT-authored sources. */
  public boolean hasStandardInput() {
    return standardInputSource.isPresent();
  }

  /** Returns the lazy host payload capability; obtaining it does not open or consume input. */
  public Optional<GridGrindPayloadOpener> standardInputSource() {
    return standardInputSource;
  }

  /** Creates execution inputs with a lazy payload and optional host-owned secret resolver. */
  public GridGrindRequestInputs(
      Path workingDirectory,
      Path tempRoot,
      GridGrindPayloadOpener standardInputSource,
      ExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(Objects.requireNonNull(standardInputSource)),
        executionGrant,
        secretResolver);
  }

  private static GridGrindPayloadOpener bytesSource(byte[] bytes, ExecutionGrant grant) {
    Objects.requireNonNull(grant)
        .requireInputSnapshotBytes(0, Objects.requireNonNull(bytes).length);
    byte[] snapshot = Objects.requireNonNull(bytes, "standardInputBytes must not be null").clone();
    return () -> new java.io.ByteArrayInputStream(snapshot);
  }

  /** Returns the trusted host grant required for this execution. */
  public ExecutionGrant executionGrant() {
    return executionGrant;
  }

  /** Returns the host-owned resolver when this execution can resolve secret references. */
  public Optional<GridGrindSecretResolver> secretResolver() {
    return secretResolver;
  }
}
