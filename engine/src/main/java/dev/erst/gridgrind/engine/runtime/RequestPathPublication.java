package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.Objects;
import java.util.Optional;
import java.util.UUID;

/** Publishes one verified staged artifact without exposing a partial final destination. */
public final class RequestPathPublication {
  private final AtomicMover atomicMover;
  private final SiblingWriter siblingWriter;

  RequestPathPublication() {
    this(RequestPathPublication::moveAtomically, RequestPathPublication::copySibling);
  }

  RequestPathPublication(AtomicMover atomicMover) {
    this(atomicMover, RequestPathPublication::copySibling);
  }

  RequestPathPublication(AtomicMover atomicMover, SiblingWriter siblingWriter) {
    this.atomicMover = Objects.requireNonNull(atomicMover, "atomicMover must not be null");
    this.siblingWriter = Objects.requireNonNull(siblingWriter, "siblingWriter must not be null");
  }

  PublicationAttempt publish(
      RequestPathBinding binding,
      Path stagedArtifact,
      WorkbookArtifactWriteDisposition disposition,
      PublicationOutcome.StagedArtifactVerification stagedArtifactVerification) {
    Objects.requireNonNull(binding, "binding must not be null");
    Objects.requireNonNull(stagedArtifact, "stagedArtifact must not be null");
    Objects.requireNonNull(disposition, "disposition must not be null");
    Objects.requireNonNull(
        stagedArtifactVerification, "stagedArtifactVerification must not be null");

    Optional<Path> siblingName = Optional.empty();
    PublicationTargetState before = new PublicationTargetUnknown();
    PublicationMoveState moveState = new PublicationMoveNotAttempted();
    try {
      binding.reverifyPublicationTarget();
      before = targetSnapshot(binding, disposition);
      WorkbookArtifactWriteDisposition atomicDisposition = atomicDisposition(disposition, before);
      RequestPathPublicationFileSupport.ArtifactDigest stagedDigest =
          RequestPathPublicationFileSupport.digest(stagedArtifact);
      Path createdSibling = nextSiblingName();
      siblingName = Optional.of(createdSibling);
      siblingWriter.write(stagedArtifact, binding, createdSibling);
      RequestPathPublicationFileSupport.ArtifactDigest siblingDigest =
          RequestPathPublicationFileSupport.digest(binding.openSiblingReadChannel(createdSibling));
      if (!siblingDigest.equals(stagedDigest)) {
        return failedAttempt(
            binding,
            stagedArtifactVerification,
            before,
            moveState,
            new IOException("private publication sibling digest differs from staged artifact"));
      }
      binding.reverifyPublicationTarget();
      moveState = new PublicationMoveAttempted(createdSibling);
      atomicMover.move(
          binding.siblingPath(createdSibling), binding.resolvedPath(), atomicDisposition);
      moveState = new PublicationMoveCompleted();
      return publishedAttempt(binding, stagedDigest, stagedArtifactVerification);
    } catch (IOException exception) {
      return failedAttempt(binding, stagedArtifactVerification, before, moveState, exception);
    } finally {
      siblingName.ifPresent(sibling -> deleteSiblingQuietly(binding, sibling));
    }
  }

  private PublicationAttempt failedAttempt(
      RequestPathBinding binding,
      PublicationOutcome.StagedArtifactVerification stagedArtifactVerification,
      PublicationTargetState before,
      PublicationMoveState moveState,
      IOException exception) {
    switch (moveState) {
      case PublicationMoveCompleted _ -> {
        return uncertainAttempt(binding, exception);
      }
      case PublicationMoveAttempted attempted -> {
        try {
          RequestPathPublicationFileSupport.ArtifactDigest stagedDigest =
              RequestPathPublicationFileSupport.digest(
                  binding.openSiblingReadChannel(attempted.siblingName()));
          Optional<RequestPathPublicationFileSupport.ArtifactDigest> current =
              currentDigest(binding);
          if (current.isPresent() && current.orElseThrow().equals(stagedDigest)) {
            return publishedAttempt(binding, stagedDigest, stagedArtifactVerification);
          }
        } catch (IOException verificationFailure) {
          exception.addSuppressed(verificationFailure);
        }
      }
      case PublicationMoveNotAttempted _ -> {}
    }
    try {
      PublicationOutcome.DestinationState state = destinationState(binding, before);
      return new FailedPublicationAttempt(new PublicationOutcome.NotPublished(state), exception);
    } catch (IOException stateFailure) {
      exception.addSuppressed(stateFailure);
      return uncertainAttempt(binding, exception);
    }
  }

  private static PublishedPublicationAttempt publishedAttempt(
      RequestPathBinding binding,
      RequestPathPublicationFileSupport.ArtifactDigest stagedDigest,
      PublicationOutcome.StagedArtifactVerification stagedArtifactVerification)
      throws IOException {
    RequestPathPublicationFileSupport.ArtifactDigest publishedDigest =
        RequestPathPublicationFileSupport.digest(binding.openPublishedReadChannel());
    if (!publishedDigest.equals(stagedDigest)) {
      throw new IOException(
          "published workbook digest differs from the verified staged artifact: "
              + binding.resolvedPath());
    }
    PublicationOutcome.DurabilityEvidence durability =
        RequestPathPublicationFileSupport.forcePublishedArtifact(binding);
    return new PublishedPublicationAttempt(
        new PublicationOutcome.Published(
            binding.resolvedPath().toString(),
            stagedDigest.sha256(),
            stagedDigest.byteSize(),
            stagedArtifactVerification,
            durability));
  }

  private static FailedPublicationAttempt uncertainAttempt(
      RequestPathBinding binding, IOException exception) {
    return new FailedPublicationAttempt(
        new PublicationOutcome.Uncertain(
            binding.resolvedPath().toString(),
            PublicationOutcome.Recovery.INSPECT_DESTINATION_DO_NOT_RETRY),
        exception);
  }

  private static PublicationTargetState targetSnapshot(
      RequestPathBinding binding, WorkbookArtifactWriteDisposition disposition) throws IOException {
    return switch (disposition) {
      case CREATE_NEW -> new PublicationTargetAbsent();
      case REPLACE_EXISTING ->
          currentDigest(binding)
              .<PublicationTargetState>map(PublicationTargetExisting::new)
              .orElseGet(PublicationTargetAbsent::new);
    };
  }

  private static PublicationOutcome.DestinationState destinationState(
      RequestPathBinding binding, PublicationTargetState before) throws IOException {
    Optional<RequestPathPublicationFileSupport.ArtifactDigest> current = currentDigest(binding);
    return switch (before) {
      case PublicationTargetUnknown _ ->
          throw new IOException(
              "publication destination state was not established before the failed publication: "
                  + binding.resolvedPath());
      case PublicationTargetAbsent _ -> {
        if (current.isEmpty()) {
          yield PublicationOutcome.DestinationState.ABSENT;
        }
        throw destinationStateChanged(binding);
      }
      case PublicationTargetExisting existing -> {
        if (current.isPresent() && existing.digest().equals(current.orElseThrow())) {
          yield PublicationOutcome.DestinationState.PRESERVED;
        }
        throw destinationStateChanged(binding);
      }
    };
  }

  private static IOException destinationStateChanged(RequestPathBinding binding) {
    return new IOException(
        "publication destination differs from its admitted state after a failed publication: "
            + binding.resolvedPath());
  }

  private static WorkbookArtifactWriteDisposition atomicDisposition(
      WorkbookArtifactWriteDisposition requestedDisposition, PublicationTargetState before) {
    if (requestedDisposition == WorkbookArtifactWriteDisposition.REPLACE_EXISTING
        && before instanceof PublicationTargetAbsent) {
      return WorkbookArtifactWriteDisposition.CREATE_NEW;
    }
    return requestedDisposition;
  }

  private static Optional<RequestPathPublicationFileSupport.ArtifactDigest> currentDigest(
      RequestPathBinding binding) throws IOException {
    try {
      return Optional.of(
          RequestPathPublicationFileSupport.digest(binding.openPublishedReadChannel()));
    } catch (NoSuchFileException ignored) {
      return Optional.empty();
    }
  }

  private static Path nextSiblingName() {
    return Path.of(".gridgrind-publish-" + UUID.randomUUID() + ".xlsx");
  }

  private static void copySibling(Path stagedArtifact, RequestPathBinding binding, Path siblingName)
      throws IOException {
    RequestPathPublicationFileSupport.copyAndForce(
        stagedArtifact,
        binding.openSiblingChannel(
            siblingName,
            java.util.Set.of(
                StandardOpenOption.CREATE_NEW,
                StandardOpenOption.WRITE,
                LinkOption.NOFOLLOW_LINKS)));
  }

  private static void moveAtomically(
      Path source, Path destination, WorkbookArtifactWriteDisposition disposition)
      throws IOException {
    if (disposition == WorkbookArtifactWriteDisposition.CREATE_NEW) {
      Files.move(source, destination, StandardCopyOption.ATOMIC_MOVE);
      return;
    }
    Files.move(
        source, destination, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
  }

  private static void deleteSiblingQuietly(RequestPathBinding binding, Path siblingName) {
    deleteSiblingQuietly(siblingName, binding::deleteSibling);
  }

  static void deleteSiblingQuietly(Path siblingName, SiblingDeletion deletion) {
    Objects.requireNonNull(siblingName, "siblingName must not be null");
    Objects.requireNonNull(deletion, "deletion must not be null");
    try {
      deletion.delete(siblingName);
    } catch (NoSuchFileException ignored) {
      // Atomic publication already moved the sibling into the final destination.
    } catch (IOException ignored) {
      // A retained final artifact state remains more actionable than private cleanup noise.
    }
  }

  /** One atomic move seam for deterministic publication-failure injection. */
  @FunctionalInterface
  interface AtomicMover {
    /** Moves one private sibling into its final destination. */
    void move(Path source, Path destination, WorkbookArtifactWriteDisposition disposition)
        throws IOException;
  }

  /** One private-sibling writer seam for deterministic interrupted-copy testing. */
  @FunctionalInterface
  interface SiblingWriter {
    /** Copies a complete staged artifact into one private sibling entry. */
    void write(Path stagedArtifact, RequestPathBinding binding, Path siblingName)
        throws IOException;
  }

  /** Deletes a private publication sibling without re-resolving its request-owned path. */
  @FunctionalInterface
  interface SiblingDeletion {
    /** Deletes the supplied private sibling name. */
    void delete(Path siblingName) throws IOException;
  }
}

/** Exact atomic-publication state known at a caught I/O boundary. */
sealed interface PublicationMoveState
    permits PublicationMoveNotAttempted, PublicationMoveAttempted, PublicationMoveCompleted {}

/** No atomic move was started. */
record PublicationMoveNotAttempted() implements PublicationMoveState {}

/** One private sibling was handed to the atomic move primitive. */
record PublicationMoveAttempted(Path siblingName) implements PublicationMoveState {
  PublicationMoveAttempted {
    Objects.requireNonNull(siblingName, "siblingName must not be null");
  }
}

/** The move returned normally but final verification did not establish publication. */
record PublicationMoveCompleted() implements PublicationMoveState {}

/** Final-destination fact established before private sibling publication begins. */
sealed interface PublicationTargetState
    permits PublicationTargetUnknown, PublicationTargetAbsent, PublicationTargetExisting {}

/** The pre-publication destination fact was not established. */
record PublicationTargetUnknown() implements PublicationTargetState {}

/** The pre-publication destination was established as absent. */
record PublicationTargetAbsent() implements PublicationTargetState {}

/** The pre-publication destination was established with these exact bytes. */
record PublicationTargetExisting(RequestPathPublicationFileSupport.ArtifactDigest digest)
    implements PublicationTargetState {
  PublicationTargetExisting {
    Objects.requireNonNull(digest, "digest must not be null");
  }
}

/** Internal outcome preserving the precise publication fact plus any primary I/O failure. */
sealed interface PublicationAttempt permits PublishedPublicationAttempt, FailedPublicationAttempt {}

/** Publication completed and final bytes were verified. */
record PublishedPublicationAttempt(PublicationOutcome.Published outcome)
    implements PublicationAttempt {
  PublishedPublicationAttempt {
    Objects.requireNonNull(outcome, "outcome must not be null");
  }
}

/** Publication did not complete or cannot be proven after a checked I/O failure. */
record FailedPublicationAttempt(PublicationOutcome outcome, IOException failure)
    implements PublicationAttempt {
  FailedPublicationAttempt {
    if (!(outcome instanceof PublicationOutcome.NotPublished)
        && !(outcome instanceof PublicationOutcome.Uncertain)) {
      throw new IllegalArgumentException("failure outcome must be NOT_PUBLISHED or UNCERTAIN");
    }
    Objects.requireNonNull(failure, "failure must not be null");
  }
}
