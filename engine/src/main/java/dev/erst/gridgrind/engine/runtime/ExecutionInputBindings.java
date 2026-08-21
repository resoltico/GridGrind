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
  private final Optional<InputResolutionOrigins> inputResolutionOrigins;
  private final Optional<RequestPathAccess> requestPathAccess;

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
    this(
        workingDirectory,
        tempRoot,
        standardInput,
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      Optional<StandardInputBinding> standardInput,
      Optional<InputResolutionFailures> inputResolutionFailures,
      Optional<InputResolutionOrigins> inputResolutionOrigins,
      Optional<RequestPathAccess> requestPathAccess) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(standardInput, "standardInput must not be null");
    Objects.requireNonNull(inputResolutionFailures, "inputResolutionFailures must not be null");
    Objects.requireNonNull(inputResolutionOrigins, "inputResolutionOrigins must not be null");
    Objects.requireNonNull(requestPathAccess, "requestPathAccess must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.tempRoot = tempRoot.toAbsolutePath().normalize();
    this.standardInput = standardInput;
    this.inputResolutionFailures = inputResolutionFailures;
    this.inputResolutionOrigins = inputResolutionOrigins;
    this.requestPathAccess = requestPathAccess;
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

  /** Returns the request-scoped no-follow filesystem capability prepared during preflight. */
  RequestPathAccess requestPathAccess() {
    return requestPathAccess.orElseThrow(
        () ->
            new IllegalStateException(
                "request-owned filesystem access requires phase-four preparation"));
  }

  boolean hasRequestPathAccess() {
    return requestPathAccess.isPresent();
  }

  ExecutionInputBindings collectingInputResolutionFailures(InputResolutionFailures failures) {
    return new ExecutionInputBindings(
        workingDirectory,
        tempRoot,
        standardInput,
        Optional.of(Objects.requireNonNull(failures, "failures must not be null")),
        inputResolutionOrigins,
        requestPathAccess);
  }

  ExecutionInputBindings withInputResolutionOrigins(InputResolutionOrigins origins) {
    return new ExecutionInputBindings(
        workingDirectory,
        tempRoot,
        standardInput,
        inputResolutionFailures,
        Optional.of(Objects.requireNonNull(origins, "origins must not be null")),
        requestPathAccess);
  }

  ExecutionInputBindings withRequestPathAccess(RequestPathAccess access) {
    if (requestPathAccess.isPresent()) {
      throw new IllegalStateException("request-owned filesystem access is already prepared");
    }
    return new ExecutionInputBindings(
        workingDirectory,
        tempRoot,
        standardInput,
        inputResolutionFailures,
        inputResolutionOrigins,
        Optional.of(Objects.requireNonNull(access, "access must not be null")));
  }

  boolean collectInputResolutionFailure(Exception failure, Object source) {
    if (inputResolutionFailures.isEmpty()) {
      return false;
    }
    inputResolutionFailures
        .orElseThrow()
        .add(
            failure,
            inputResolutionOrigins.flatMap(origins -> origins.locationFor(source)),
            source);
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
