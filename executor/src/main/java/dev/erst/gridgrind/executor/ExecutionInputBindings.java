package dev.erst.gridgrind.executor;

import java.nio.file.Path;
import java.util.Objects;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class ExecutionInputBindings {
  private final Path workingDirectory;
  private final StandardInputBinding standardInput;

  /** Creates bindings from one working directory plus optional stdin bytes. */
  public ExecutionInputBindings(Path workingDirectory, byte[] standardInputBytes) {
    this(
        workingDirectory,
        standardInputBytes == null ? null : new StandardInputBinding(standardInputBytes));
  }

  /** Creates bindings from one working directory plus one optional stdin binding. */
  public ExecutionInputBindings(Path workingDirectory, StandardInputBinding standardInput) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.standardInput = standardInput;
  }

  /** Returns bindings rooted at the current process working directory with no bound stdin data. */
  public static ExecutionInputBindings processDefault() {
    return new ExecutionInputBindings(Path.of(""), (StandardInputBinding) null);
  }

  /** Returns the normalized working directory used to resolve relative authored input paths. */
  public Path workingDirectory() {
    return workingDirectory;
  }

  /** Returns true when stdin bytes are available to STANDARD_INPUT-authored sources. */
  public boolean hasStandardInput() {
    return standardInput != null;
  }

  /** Returns one defensive copy of the bound stdin bytes, or {@code null} when none are bound. */
  public byte[] standardInputBytes() {
    return standardInput == null ? null : standardInput.bytes();
  }

  /** Immutable standard-input byte payload bound to one execution. */
  public static final class StandardInputBinding {
    private final byte[] bytes;

    /** Creates one immutable stdin binding from the provided bytes. */
    public StandardInputBinding(byte[] bytes) {
      Objects.requireNonNull(bytes, "bytes must not be null");
      this.bytes = bytes.clone();
    }

    /** Returns one defensive copy of the bound stdin bytes. */
    public byte[] bytes() {
      return bytes.clone();
    }
  }
}
