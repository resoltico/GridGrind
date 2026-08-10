package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class ExecutionInputBindings {
  private final Path workingDirectory;
  private final Path tempRoot;
  private final Optional<StandardInputBinding> standardInput;
  private final Optional<InputResolutionFailures> inputResolutionFailures;

  /** Creates bindings from one working directory and one explicit temp root. */
  public ExecutionInputBindings(Path workingDirectory, Path tempRoot) {
    this(workingDirectory, tempRoot, Optional.empty());
  }

  /** Creates bindings from one working directory, temp root, and one stdin payload. */
  public ExecutionInputBindings(Path workingDirectory, Path tempRoot, byte[] standardInputBytes) {
    this(workingDirectory, tempRoot, new StandardInputBinding(standardInputBytes));
  }

  /** Creates bindings from one working directory, temp root, and one stdin binding. */
  public ExecutionInputBindings(
      Path workingDirectory, Path tempRoot, StandardInputBinding standardInput) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(Objects.requireNonNull(standardInput, "standardInput")));
  }

  private ExecutionInputBindings(
      Path workingDirectory, Path tempRoot, Optional<StandardInputBinding> standardInput) {
    this(workingDirectory, tempRoot, standardInput, Optional.empty());
  }

  private ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      Optional<StandardInputBinding> standardInput,
      Optional<InputResolutionFailures> inputResolutionFailures) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(standardInput, "standardInput must not be null");
    Objects.requireNonNull(inputResolutionFailures, "inputResolutionFailures must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.tempRoot = tempRoot.toAbsolutePath().normalize();
    this.standardInput = standardInput;
    this.inputResolutionFailures = inputResolutionFailures;
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
    return standardInput.isPresent();
  }

  /** Returns one defensive copy of the bound stdin bytes when a stdin binding is present. */
  public Optional<byte[]> standardInputBytes() {
    return standardInput.map(StandardInputBinding::bytes);
  }

  /** Returns one temp-file factory rooted at this execution's explicit temp directory. */
  TempFileFactory tempFileFactory() {
    WorkbookTempFileFactory workbookTempFileFactory = WorkbookTempFileFactory.rooted(tempRoot);
    return workbookTempFileFactory::createTempFile;
  }

  ExecutionInputBindings collectingInputResolutionFailures(InputResolutionFailures failures) {
    return new ExecutionInputBindings(
        workingDirectory,
        tempRoot,
        standardInput,
        Optional.of(Objects.requireNonNull(failures, "failures must not be null")));
  }

  boolean collectInputResolutionFailure(Exception failure) {
    if (inputResolutionFailures.isEmpty()) {
      return false;
    }
    inputResolutionFailures.orElseThrow().add(failure);
    return true;
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
