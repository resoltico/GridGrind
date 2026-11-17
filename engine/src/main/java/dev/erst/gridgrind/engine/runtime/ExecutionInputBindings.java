package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.engine.api.GridGrindSecretResolver;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Transport-neutral authored-input bindings supplied alongside one executed workbook plan. */
public final class ExecutionInputBindings {
  private final ExecutionSession session;
  private final MediaMaterializationUse mediaUse;
  private final Path workingDirectory;
  private final Path tempRoot;
  private final Optional<StandardInputBinding> standardInput;
  private final ExecutionGrant executionGrant;
  private final Optional<GridGrindSecretResolver> secretResolver;
  private final Optional<InputResolutionFailures> inputResolutionFailures;
  private final Optional<InputResolutionOrigins> inputResolutionOrigins;
  private final Optional<RequestPathAccess> requestPathAccess;

  /** Creates bindings from one working directory, temp root, and explicit host grant. */
  public ExecutionInputBindings(
      Path workingDirectory, Path tempRoot, ExecutionGrant executionGrant) {
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
      ExecutionGrant executionGrant) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(new StandardInputBinding(standardInputBytes, executionGrant)),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.empty());
  }

  /** Creates stdin bindings with a host-owned resolver for plan secret references. */
  public ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      byte[] standardInputBytes,
      ExecutionGrant executionGrant,
      GridGrindSecretResolver secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        Optional.of(new StandardInputBinding(standardInputBytes, executionGrant)),
        Objects.requireNonNull(executionGrant, "executionGrant"),
        Optional.of(Objects.requireNonNull(secretResolver, "secretResolver")));
  }

  /** Creates bindings from one working directory, temp root, stdin binding, and host grant. */
  public ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      StandardInputBinding standardInput,
      ExecutionGrant executionGrant) {
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
      ExecutionGrant executionGrant,
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
      ExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver) {
    this(
        workingDirectory,
        tempRoot,
        standardInput,
        executionGrant,
        secretResolver,
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        new ExecutionSession(tempRoot, executionGrant));
  }

  private ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      Optional<StandardInputBinding> standardInput,
      ExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver,
      Optional<InputResolutionFailures> inputResolutionFailures,
      Optional<InputResolutionOrigins> inputResolutionOrigins,
      Optional<RequestPathAccess> requestPathAccess,
      ExecutionSession session) {
    this(
        workingDirectory,
        tempRoot,
        standardInput,
        executionGrant,
        secretResolver,
        inputResolutionFailures,
        inputResolutionOrigins,
        requestPathAccess,
        session,
        new MediaMaterializationUse(session.media));
  }

  private ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      Optional<StandardInputBinding> standardInput,
      ExecutionGrant executionGrant,
      Optional<GridGrindSecretResolver> secretResolver,
      Optional<InputResolutionFailures> inputResolutionFailures,
      Optional<InputResolutionOrigins> inputResolutionOrigins,
      Optional<RequestPathAccess> requestPathAccess,
      ExecutionSession session,
      MediaMaterializationUse mediaUse) {
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(standardInput, "standardInput must not be null");
    Objects.requireNonNull(executionGrant, "executionGrant must not be null");
    Objects.requireNonNull(secretResolver, "secretResolver must not be null");
    Objects.requireNonNull(inputResolutionFailures, "inputResolutionFailures must not be null");
    Objects.requireNonNull(inputResolutionOrigins, "inputResolutionOrigins must not be null");
    Objects.requireNonNull(requestPathAccess, "requestPathAccess must not be null");
    this.session = Objects.requireNonNull(session);
    this.mediaUse = Objects.requireNonNull(mediaUse);
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

  /** Returns a read-only snapshot view with an independent cursor for internal decoding. */
  Optional<java.nio.ByteBuffer> standardInputSnapshot() throws java.io.IOException {
    requireStandardInputAuthority(executionGrant);
    if (standardInput.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(session.inputs.standardInput(standardInput.orElseThrow().opener));
  }

  Optional<java.nio.ByteBuffer> standardInputSnapshotWithin(InputSnapshotLimit limit)
      throws java.io.IOException {
    requireStandardInputAuthority(executionGrant);
    if (standardInput.isEmpty()) return Optional.empty();
    return Optional.of(
        session.inputs.standardInputWithin(standardInput.orElseThrow().opener, limit));
  }

  dev.erst.gridgrind.engine.media.MediaAdmissionBudget mediaBudget() {
    return session.media;
  }

  MediaMaterializationUse mediaUse() {
    return mediaUse;
  }

  ExecutionInputBindings withNewMediaUse() {
    return new ExecutionInputBindings(
        workingDirectory,
        tempRoot,
        standardInput,
        executionGrant,
        secretResolver,
        inputResolutionFailures,
        inputResolutionOrigins,
        requestPathAccess,
        session,
        new MediaMaterializationUse(session.media));
  }

  ExecutionSession session() {
    return session;
  }

  dev.erst.gridgrind.core.grant.AcceptancePolicy acceptancePolicy() {
    return session.acceptancePolicy;
  }

  void resolvedAcceptancePolicy(dev.erst.gridgrind.core.grant.AcceptancePolicy policy) {
    session.acceptancePolicy = Objects.requireNonNull(policy);
  }

  void recordExecutionStrategy(
      dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.ExecutionStrategy value) {
    session.recordExecutionStrategy(value);
  }

  java.util.Optional<dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.ExecutionStrategy>
      executionStrategy() {
    return session.executionStrategy();
  }

  PublicationFacts publicationFacts() {
    return session.publication;
  }

  InputSnapshots snapshots() {
    return session.inputs;
  }

  /** Creates host bindings without invoking the lazy payload capability. */
  public ExecutionInputBindings(
      Path workingDirectory,
      Path tempRoot,
      ExecutionGrant grant,
      Optional<dev.erst.gridgrind.engine.api.GridGrindPayloadOpener> source,
      Optional<GridGrindSecretResolver> resolver) {
    this(workingDirectory, tempRoot, source.map(StandardInputBinding::new), grant, resolver);
  }

  /** Returns the trusted host grant required for this execution. */
  ExecutionGrant executionGrant() {
    return executionGrant;
  }

  String resolveSecret(SecretReference reference) {
    Objects.requireNonNull(reference, "reference must not be null");
    ExecutionGrant grant = executionGrant;
    if (!grant.allowedSecretReferences().contains(reference)) {
      throw new ExecutionAuthorityDeniedException(
          "host grant does not permit secret reference " + reference.id());
    }
    GridGrindSecretResolver resolver =
        secretResolver.orElseThrow(() ->
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

  private void requireStandardInputAuthority(ExecutionGrant grant) {
    if (standardInput.isEmpty()) {
      return;
    }
    boolean allowed =
        grant.readableResources().stream()
            .anyMatch(ExecutionGrant.ReadAuthority.Stdin.class::isInstance);
    if (!allowed) {
      throw new ExecutionAuthorityDeniedException(
          "host grant does not permit reading the standard-input payload");
    }
  }

  /** Returns one temp-file factory rooted at this execution's explicit temp directory. */
  TempFileFactory tempFileFactory() {
    return session::createTempFile;
  }

  ExecutionInputBindings forRun() {
    return new ExecutionInputBindings(
        workingDirectory, tempRoot, standardInput, executionGrant, secretResolver);
  }

  void own(RequestPathAccess.Cleanup cleanup) {
    session.own(cleanup);
  }

  void closeRunResources() throws java.io.IOException {
    session.close();
  }

  /** Returns the request-scoped no-follow filesystem capability prepared during preflight. */
  RequestPathAccess requestPathAccess() {
    return requestPathAccess.orElseThrow(() ->
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
        requestPathAccess,
        session,
        mediaUse);
  }

  ExecutionInputBindings withInputResolutionScope(String path) {
    return inputResolutionOrigins
        .map(origins -> withInputResolutionOrigins(origins.scoped(path)))
        .orElse(this);
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
        requestPathAccess,
        session,
        mediaUse);
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
        Optional.of(Objects.requireNonNull(access, "access must not be null")),
        session,
        mediaUse);
  }

  Optional<dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation> inputOrigin(
      Object source) {
    return inputResolutionOrigins.flatMap(origins -> origins.locationFor(source));
  }

  Optional<dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation>
      lastInputOrigin(Object source) {
    return inputResolutionOrigins.flatMap(origins -> origins.lastLocationFor(source));
  }

  boolean collectInputResolutionFailure(
      Exception failure,
      Object source,
      Optional<dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation>
          location) {
    if (inputResolutionFailures.isEmpty()) {
      return false;
    }
    inputResolutionFailures.orElseThrow().add(failure, location, source);
    return true;
  }

  /** Immutable standard-input byte payload bound to one execution. */
  public static final class StandardInputBinding {
    private final dev.erst.gridgrind.engine.api.GridGrindPayloadOpener opener;

    /** Creates one immutable stdin binding from the provided bytes. */
    public StandardInputBinding(byte[] bytes, ExecutionGrant grant) {
      Objects.requireNonNull(grant)
          .requireInputSnapshotBytes(0, Objects.requireNonNull(bytes).length);
      Objects.requireNonNull(bytes, "bytes must not be null");
      byte[] snapshot = bytes.clone();
      this.opener = () -> new java.io.ByteArrayInputStream(snapshot);
    }

    private StandardInputBinding(dev.erst.gridgrind.engine.api.GridGrindPayloadOpener opener) {
      this.opener = Objects.requireNonNull(opener);
    }
  }
}
