package dev.erst.gridgrind.engine.api;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class GridGrindRequestInputs {
  private final Path workingDirectory;
  private final Optional<byte[]> standardInputBytes;

  /** Creates execution inputs from one working directory with no bound stdin bytes. */
  public GridGrindRequestInputs(Path workingDirectory) {
    this(workingDirectory, Optional.empty());
  }

  /** Creates execution inputs from one working directory plus one bound stdin payload. */
  public GridGrindRequestInputs(Path workingDirectory, byte[] standardInputBytes) {
    this(
        workingDirectory,
        Optional.of(Objects.requireNonNull(standardInputBytes, "standardInputBytes").clone()));
  }

  private GridGrindRequestInputs(Path workingDirectory, Optional<byte[]> standardInputBytes) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(standardInputBytes, "standardInputBytes must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.standardInputBytes = standardInputBytes.map(byte[]::clone);
  }

  /** Returns inputs rooted at the current process working directory with no bound stdin data. */
  public static GridGrindRequestInputs processDefault() {
    return new GridGrindRequestInputs(Path.of(""));
  }

  /** Returns the normalized working directory used to resolve relative authored input paths. */
  public Path workingDirectory() {
    return workingDirectory;
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
