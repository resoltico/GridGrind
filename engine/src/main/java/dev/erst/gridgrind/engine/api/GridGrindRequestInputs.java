package dev.erst.gridgrind.engine.api;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class GridGrindRequestInputs {
  private final Path workingDirectory;
  private final Path tempRoot;
  private final Optional<byte[]> standardInputBytes;

  /** Creates execution inputs from one working directory plus one explicit temp root. */
  public GridGrindRequestInputs(Path workingDirectory, Path tempRoot) {
    this(workingDirectory, tempRoot, Optional.empty());
  }

  /** Creates execution inputs from one working directory, temp root, and stdin payload. */
  public GridGrindRequestInputs(Path workingDirectory, Path tempRoot, byte[] standardInputBytes) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(Objects.requireNonNull(standardInputBytes, "standardInputBytes").clone()));
  }

  private GridGrindRequestInputs(
      Path workingDirectory, Path tempRoot, Optional<byte[]> standardInputBytes) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(standardInputBytes, "standardInputBytes must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.tempRoot = tempRoot.toAbsolutePath().normalize();
    this.standardInputBytes = standardInputBytes.map(byte[]::clone);
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
}
