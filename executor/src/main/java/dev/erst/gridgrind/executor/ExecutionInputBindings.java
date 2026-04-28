package dev.erst.gridgrind.executor;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class ExecutionInputBindings {
  private final Path workingDirectory;
  private final Optional<StandardInputBinding> standardInput;

  /** Creates bindings from one working directory with no bound stdin bytes. */
  public ExecutionInputBindings(Path workingDirectory) {
    this(workingDirectory, Optional.empty());
  }

  /** Creates bindings from one working directory plus one bound stdin payload. */
  public ExecutionInputBindings(Path workingDirectory, byte[] standardInputBytes) {
    this(workingDirectory, new StandardInputBinding(standardInputBytes));
  }

  /** Creates bindings from one working directory plus one explicit stdin binding. */
  public ExecutionInputBindings(Path workingDirectory, StandardInputBinding standardInput) {
    this(workingDirectory, Optional.of(Objects.requireNonNull(standardInput, "standardInput")));
  }

  private ExecutionInputBindings(
      Path workingDirectory, Optional<StandardInputBinding> standardInput) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(standardInput, "standardInput must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.standardInput = standardInput;
  }

  /** Returns bindings rooted at the current process working directory with no bound stdin data. */
  public static ExecutionInputBindings processDefault() {
    return new ExecutionInputBindings(Path.of(""));
  }

  /** Returns the normalized working directory used to resolve relative authored input paths. */
  public Path workingDirectory() {
    return workingDirectory;
  }

  /** Returns true when stdin bytes are available to STANDARD_INPUT-authored sources. */
  public boolean hasStandardInput() {
    return standardInput.isPresent();
  }

  /** Returns one defensive copy of the bound stdin bytes when a stdin binding is present. */
  public Optional<byte[]> standardInputBytes() {
    return standardInput.map(StandardInputBinding::bytes);
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
