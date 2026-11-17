package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.FileAlreadyExistsException;
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
  private final PublicationFileOperations.FileForcer fileForcer;
  private final PublicationFacts facts;
  private final PublicationCommitter committer;
  private final SiblingWriter siblingWriter;

  RequestPathPublication() {
    this(RequestPathPublication::commitAtomically, RequestPathPublication::copySibling);
  }

  RequestPathPublication(PublicationCommitter committer) {
    this(committer, RequestPathPublication::copySibling);
  }

  RequestPathPublication(PublicationFacts facts) {
    this(facts, RequestPathPublication::commitAtomically, RequestPathPublication::copySibling);
  }

  RequestPathPublication(PublicationCommitter committer, SiblingWriter siblingWriter) {
    this(new PublicationFacts(), committer, siblingWriter);
  }

  RequestPathPublication(
      PublicationCommitter committer,
      SiblingWriter siblingWriter,
      PublicationFileOperations.FileForcer fileForcer) {
    this(new PublicationFacts(), committer, siblingWriter, fileForcer);
  }

  private RequestPathPublication(
      PublicationFacts facts, PublicationCommitter committer, SiblingWriter siblingWriter) {
    this(facts, committer, siblingWriter, channel -> channel.force(true));
  }

  private RequestPathPublication(
      PublicationFacts facts,
      PublicationCommitter committer,
      SiblingWriter siblingWriter,
      PublicationFileOperations.FileForcer fileForcer) {
    this.fileForcer = Objects.requireNonNull(fileForcer);
    this.facts = Objects.requireNonNull(facts);
    this.committer = Objects.requireNonNull(committer, "committer must not be null");
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
    PublicationCommitState commitState = new PublicationCommitNotAttempted();
    try {
      binding.reverifyPublicationTarget();
      WorkbookArtifactWriteDisposition atomicDisposition = disposition;
      PublicationFileOperations.ArtifactDigest stagedDigest =
          PublicationFileOperations.digest(stagedArtifact);
      Path createdSibling = nextSiblingName();
      siblingName = Optional.of(createdSibling);
      siblingWriter.write(stagedArtifact, binding, createdSibling);
      PublicationFileOperations.ArtifactDigest siblingDigest =
          PublicationFileOperations.digest(binding.openSiblingReadChannel(createdSibling));
      if (!siblingDigest.equals(stagedDigest)) {
        return failedAttempt(
            binding,
            commitState,
            new IOException("private publication sibling digest differs from staged artifact"));
      }
      binding.reverifyPublicationTarget();
      if (atomicDisposition == WorkbookArtifactWriteDisposition.CREATE_NEW) {
        requireQualifiedNoClobberProvider(binding.siblingPath(createdSibling));
      }
      commitState =
          new PublicationCommitAttempted(
              atomicDisposition == WorkbookArtifactWriteDisposition.CREATE_NEW);
      facts.attempting(binding.resolvedPath().toString());
      committer.commit(
          binding.siblingPath(createdSibling), binding.resolvedPath(), atomicDisposition);
      facts.returnedSuccessfully(binding.resolvedPath().toString());
      commitState = new PublicationCommitCompleted();
      var published = publishedAttempt(binding, stagedDigest, stagedArtifactVerification);
      facts.established(published.outcome());
      return published;
    } catch (IOException exception) {
      return failedAttempt(binding, commitState, exception);
    } finally {
      siblingName.ifPresent(sibling -> cleanupSibling(binding, sibling));
    }
  }

  private PublicationAttempt failedAttempt(
      RequestPathBinding binding, PublicationCommitState commitState, IOException exception) {
    PublicationAttempt attempt =
        switch (commitState) {
          case PublicationCommitNotAttempted _ ->
              new FailedPublicationAttempt(new PublicationOutcome.NotAttempted(), exception);
          case PublicationCommitCompleted _ ->
              uncertainAttempt(
                  binding, exception, PublicationOutcome.CommitObservation.RETURNED_SUCCESSFULLY);
          case PublicationCommitAttempted attempted -> {
            if (attempted.noClobber() && exception instanceof FileAlreadyExistsException) {
              yield new FailedPublicationAttempt(
                  new PublicationOutcome.NotPublished(
                      PublicationOutcome.DestinationState.COLLISION_OBSERVED),
                  exception);
            }
            yield uncertainAttempt(
                binding, exception, PublicationOutcome.CommitObservation.NOT_ESTABLISHED);
          }
        };
    facts.established(((FailedPublicationAttempt) attempt).outcome());
    return attempt;
  }

  private PublishedPublicationAttempt publishedAttempt(
      RequestPathBinding binding,
      PublicationFileOperations.ArtifactDigest stagedDigest,
      PublicationOutcome.StagedArtifactVerification stagedArtifactVerification)
      throws IOException {
    PublicationFileOperations.ArtifactDigest publishedDigest =
        PublicationFileOperations.digest(binding.openPublishedReadChannel());
    if (!publishedDigest.equals(stagedDigest)) {
      throw new IOException(
          "published workbook digest differs from the verified staged artifact: "
              + binding.resolvedPath());
    }
    PublicationOutcome.DurabilityEvidence durability =
        PublicationFileOperations.forcePublishedArtifact(
            binding, fileForcer, PublicationFileOperations::forceDirectory);
    return new PublishedPublicationAttempt(
        new PublicationOutcome.Published(
            binding.resolvedPath().toString(),
            stagedDigest.sha256(),
            stagedDigest.byteSize(),
            stagedArtifactVerification,
            durability));
  }

  private static FailedPublicationAttempt uncertainAttempt(
      RequestPathBinding binding,
      IOException exception,
      PublicationOutcome.CommitObservation observation) {
    return new FailedPublicationAttempt(
        new PublicationOutcome.Uncertain(
            binding.resolvedPath().toString(),
            PublicationOutcome.Recovery.INSPECT_DESTINATION_DO_NOT_RETRY,
            observation),
        exception);
  }

  private static Path nextSiblingName() {
    return Path.of(".gridgrind-publish-" + UUID.randomUUID() + ".xlsx");
  }

  private static void copySibling(Path stagedArtifact, RequestPathBinding binding, Path siblingName)
      throws IOException {
    PublicationFileOperations.copyAndForce(
        stagedArtifact,
        binding.openSiblingChannel(
            siblingName,
            java.util.Set.of(
                StandardOpenOption.CREATE_NEW,
                StandardOpenOption.WRITE,
                LinkOption.NOFOLLOW_LINKS)));
  }

  static void commitAtomically(
      Path source, Path destination, WorkbookArtifactWriteDisposition disposition)
      throws IOException {
    if (disposition == WorkbookArtifactWriteDisposition.CREATE_NEW) {
      Files.createLink(destination, source);
      return;
    }
    Files.move(
        source, destination, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
  }

  private static void requireQualifiedNoClobberProvider(Path source) throws IOException {
    String type = Files.getFileStore(source).type().toLowerCase(java.util.Locale.ROOT);
    String provider = source.getFileSystem().provider().getClass().getName();
    if (!java.util.Set.of(
                "sun.nio.fs.LinuxFileSystemProvider", "sun.nio.fs.MacOSXFileSystemProvider")
            .contains(provider)
        || !source
            .getFileSystem()
            .provider()
            .equals(java.nio.file.FileSystems.getDefault().provider())
        || !java.util.Set.of("apfs", "ext4", "xfs", "overlay", "overlayfs", "tmpfs", "fakeowner")
            .contains(type)) {
      throw new AtomicMoveNotSupportedException(
          source.toString(),
          null,
          "atomic no-clobber publication is not qualified for this filesystem provider");
    }
  }

  private void cleanupSibling(RequestPathBinding binding, Path siblingName) {
    try {
      binding.deleteSibling(siblingName);
    } catch (NoSuchFileException ignored) {
      // Replacement publication transferred the sibling entry.
    } catch (IOException failure) {
      facts.cleanupFailed(failure);
    }
  }

  /** One atomic commit seam for deterministic publication-failure injection. */
  @FunctionalInterface
  interface PublicationCommitter {
    /** Commits one private sibling to its final destination. */
    void commit(Path source, Path destination, WorkbookArtifactWriteDisposition disposition)
        throws IOException;
  }

  /** One private-sibling writer seam for deterministic interrupted-copy testing. */
  @FunctionalInterface
  interface SiblingWriter {
    /** Copies a complete staged artifact into one private sibling entry. */
    void write(Path stagedArtifact, RequestPathBinding binding, Path siblingName)
        throws IOException;
  }
}

/** Exact publication state known at a caught I/O boundary. */
sealed interface PublicationCommitState
    permits PublicationCommitNotAttempted, PublicationCommitAttempted, PublicationCommitCompleted {}

/** The final mutation primitive was not invoked. */
record PublicationCommitNotAttempted() implements PublicationCommitState {}

/** One private sibling was handed to the commit primitive. */
record PublicationCommitAttempted(boolean noClobber) implements PublicationCommitState {}

/** The commit returned normally but final verification did not establish publication. */
record PublicationCommitCompleted() implements PublicationCommitState {}

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
    if (!(outcome instanceof PublicationOutcome.NotAttempted)
        && !(outcome instanceof PublicationOutcome.NotPublished)
        && !(outcome instanceof PublicationOutcome.Uncertain)) {
      throw new IllegalArgumentException(
          "failure outcome must be NOT_ATTEMPTED, NOT_PUBLISHED or UNCERTAIN");
    }
    Objects.requireNonNull(failure, "failure must not be null");
  }
}
