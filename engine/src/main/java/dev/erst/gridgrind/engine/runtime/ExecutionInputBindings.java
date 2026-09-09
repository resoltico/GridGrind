package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindSecretResolver;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class ExecutionInputBindings {
  private final Path workingDirectory;
  private final Path tempRoot;
  private final Optional<StandardInputBinding> standardInput;
  private final GridGrindExecutionGrant executionGrant;
  private final Optional<GridGrindSecretResolver> secretResolver;
  private final Optional<InputResolutionFailures> inputResolutionFailures;
  private final Optional<InputResolutionOrigins> inputResolutionOrigins;
  private final Optional<RequestPathAccess> requestPathAccess;

  /** Creates bindings from one working directory, temp root, and explicit host grant. */
  public ExecutionInputBindings(
      Path workingDirectory, Path tempRoot, GridGrindExecutionGrant executionGrant) {
    this(
        workingDirectory,
        tempRoot,
        Optional.empty(),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.empty());
  }

  /** Creates bindings from one working directory, temp root, stdin payload, and host grant. */
  public ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      byte[] standardInputBytes,
      GridGrindExecutionGrant executionGrant) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(new StandardInputBinding(standardInputBytes)),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.empty());
  }

  /** Creates stdin bindings with a host-owned resolver for plan secret references. */
  public ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      byte[] standardInputBytes,
      GridGrindExecutionGrant executionGrant,
      GridGrindSecretResolver secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(new StandardInputBinding(standardInputBytes)),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.of(Objects.requireNonNull(secretResolver, "secretResolver")));
  }

  /** Creates bindings from one working directory, temp root, stdin binding, and host grant. */
  public ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      StandardInputBinding standardInput,
      GridGrindExecutionGrant executionGrant) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(Objects.requireNonNull(standardInput, "standardInput")),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.empty());
  }

  /** Creates bindings with a host-owned resolver for plan secret references. */
  public ExecutionInputBindings(
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

  private ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      Optional<StandardInputBinding> standardInput,
      GridGrindExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        standardInput,
        executionGrant,
        secretResolver,
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      Optional<StandardInputBinding> standardInput,
      GridGrindExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver,
      Optional<InputResolutionFailures> inputResolutionFailures,
      Optional<InputResolutionOrigins> inputResolutionOrigins,
      Optional<RequestPathAccess> requestPathAccess) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(standardInput, "standardInput must not be null");
    Objects.requireNonNull(executionGrant, "executionGrant must not be null");
    Objects.requireNonNull(secretResolver, "secretResolver must not be null");
    Objects.requireNonNull(inputResolutionFailures, "inputResolutionFailures must not be null");
    Objects.requireNonNull(inputResolutionOrigins, "inputResolutionOrigins must not be null");
    Objects.requireNonNull(requestPathAccess, "requestPathAccess must not be null");
    this.workingDirectory = workingDirectory.toAbsolutePath().normalize();
    this.tempRoot = tempRoot.toAbsolutePath().normalize();
    this.standardInput = standardInput;
    this.executionGrant = executionGrant;
    this.secretResolver = secretResolver;
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
    requireStandardInputAuthority(executionGrant);
    return standardInput.map(StandardInputBinding::bytes);
  }

  /** Returns the trusted host grant required for this execution. */
  GridGrindExecutionGrant executionGrant() {
    return executionGrant;
  }

  String resolveSecret(SecretReference reference) {
    Objects.requireNonNull(reference, "reference must not be null");
    GridGrindExecutionGrant.Bounded grant = (GridGrindExecutionGrant.Bounded) executionGrant;
    if (!grant.allowedSecretReferences().contains(reference)) {
      throw new ExecutionAuthorityDeniedException(
          "host grant does not permit secret reference " + reference.id());
    }
    GridGrindSecretResolver resolver =
        secretResolver.orElseThrow(
            () ->
                new ExecutionAuthorityDeniedException(
                    "host did not supply a resolver for secret reference " + reference.id()));
    char[] material =
        Objects.requireNonNull(resolver.resolve(reference), "secret resolver returned null");
    try {
      return new String(material);
    } finally {
      java.util.Arrays.fill(material, '\0');
    }
  }

  private void requireStandardInputAuthority(GridGrindExecutionGrant grant) {
    if (standardInput.isEmpty()) {
      return;
    }
    boolean allowed =
        switch (grant) {
          case GridGrindExecutionGrant.Bounded bounded ->
              bounded.readableResources().stream()
                  .anyMatch(GridGrindExecutionGrant.ReadAuthority.StandardInput.class::isInstance);
        };
    if (!allowed) {
      throw new ExecutionAuthorityDeniedException(
          "host grant does not permit reading the standard-input payload");
    }
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
        executionGrant,
        secretResolver,
        Optional.of(Objects.requireNonNull(failures, "failures must not be null")),
        inputResolutionOrigins,
        requestPathAccess);
  }

  ExecutionInputBindings withInputResolutionOrigins(InputResolutionOrigins origins) {
    return new ExecutionInputBindings(
        workingDirectory,
        tempRoot,
        standardInput,
        executionGrant,
        secretResolver,
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
        executionGrant,
        secretResolver,
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
