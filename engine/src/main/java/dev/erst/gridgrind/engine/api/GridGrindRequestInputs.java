package dev.erst.gridgrind.engine.api;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class GridGrindRequestInputs {
  private final Path workingDirectory;
  private final Path tempRoot;
  private final Optional<byte[]> standardInputBytes;
  private final GridGrindExecutionGrant executionGrant;
  private final Optional<GridGrindSecretResolver> secretResolver;

  /** Creates execution inputs from one working directory, temp root, and explicit host grant. */
  public GridGrindRequestInputs(
      Path workingDirectory, Path tempRoot, GridGrindExecutionGrant executionGrant) {
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
      GridGrindExecutionGrant executionGrant) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(Objects.requireNonNull(standardInputBytes, "standardInputBytes").clone()),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.empty());
  }

  /** Creates execution inputs with standard input, host authority, and a secret resolver. */
  public GridGrindRequestInputs(
      Path workingDirectory,
      Path tempRoot,
      byte[] standardInputBytes,
      GridGrindExecutionGrant executionGrant,
      GridGrindSecretResolver secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(Objects.requireNonNull(standardInputBytes, "standardInputBytes").clone()),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.of(Objects.requireNonNull(secretResolver, "secretResolver")));
  }

  /** Creates execution inputs with a host-owned secret resolver. */
  public GridGrindRequestInputs(
      Path workingDirectory,
      Path tempRoot,
      GridGrindExecutionGrant executionGrant,
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
      Optional<byte[]> standardInputBytes,
      GridGrindExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(standardInputBytes, "standardInputBytes must not be null");
    Objects.requireNonNull(executionGrant, "executionGrant must not be null");
    Objects.requireNonNull(secretResolver, "secretResolver must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.tempRoot = tempRoot.toAbsolutePath().normalize();
    this.standardInputBytes = standardInputBytes.map(byte[]::clone);
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
    return standardInputBytes.isPresent();
  }

  /** Returns one defensive copy of the bound stdin bytes when a stdin binding is present. */
  public Optional<byte[]> standardInputBytes() {
    return standardInputBytes.map(byte[]::clone);
  }

  /** Returns the trusted host grant required for this execution. */
  public GridGrindExecutionGrant executionGrant() {
    return executionGrant;
  }

  /** Returns the host-owned resolver when this execution can resolve secret references. */
  public Optional<GridGrindSecretResolver> secretResolver() {
    return secretResolver;
  }
}
