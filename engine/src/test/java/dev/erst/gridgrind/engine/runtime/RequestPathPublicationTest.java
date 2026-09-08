package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.assertNoPrivatePublicationSibling;
import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.moveAtomically;
import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.verifiedStage;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises truthful final-destination outcomes for descriptor-bound workbook publication. */
class RequestPathPublicationTest {
  @TempDir Path root;

  @Test
  void publishesVerifiedReplacementArtifactOverAnExistingDestination() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = Files.write(root.resolve("output.xlsx"), new byte[] {1, 2, 3, 4});

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      PublishedPublicationAttempt attempt =
          assertInstanceOf(
              PublishedPublicationAttempt.class,
              new RequestPathPublication()
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      verifiedStage()));

      assertEquals(destination.toString(), attempt.outcome().executionPath());
    }

    assertArrayEquals(new byte[] {6, 7, 8}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void publishesVerifiedReplacementArtifactWhenTheDestinationIsAbsent() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      PublishedPublicationAttempt attempt =
          assertInstanceOf(
              PublishedPublicationAttempt.class,
              new RequestPathPublication()
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      verifiedStage()));

      assertEquals(destination.toString(), attempt.outcome().executionPath());
    }

    assertArrayEquals(new byte[] {6, 7, 8}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void preservesOverwriteDestinationWhenAtomicPublicationDoesNotBegin() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = Files.write(root.resolve("source.xlsx"), new byte[] {1, 2, 3, 4});

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("source.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      (source, target, disposition) -> {
                        throw new IOException("synthetic atomic move failure");
                      })
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      verifiedStage()));

      PublicationOutcome.NotPublished outcome =
          assertInstanceOf(PublicationOutcome.NotPublished.class, attempt.outcome());
      assertEquals(PublicationOutcome.DestinationState.PRESERVED, outcome.destinationState());
    }

    assertArrayEquals(new byte[] {1, 2, 3, 4}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void preservesOverwriteDestinationWhenPrivateSiblingCopyFailsAfterOneByte() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = Files.write(root.resolve("source.xlsx"), new byte[] {1, 2, 3, 4});

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("source.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      RequestPathPublicationTestSupport::moveAtomically,
                      RequestPathPublicationTestSupport::writeOneByteThenFail)
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      verifiedStage()));

      PublicationOutcome.NotPublished outcome =
          assertInstanceOf(PublicationOutcome.NotPublished.class, attempt.outcome());
      assertEquals(PublicationOutcome.DestinationState.PRESERVED, outcome.destinationState());
    }

    assertArrayEquals(new byte[] {1, 2, 3, 4}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void reportsUncertainWhenAnEstablishedReplacementDestinationDisappearsBeforePublication()
      throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = Files.write(root.resolve("source.xlsx"), new byte[] {1, 2, 3, 4});

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("source.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      RequestPathPublicationTestSupport::moveAtomically,
                      (stagedArtifact, targetBinding, siblingName) -> {
                        Files.delete(targetBinding.resolvedPath());
                        throw new IOException("synthetic external destination removal");
                      })
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      verifiedStage()));

      assertInstanceOf(PublicationOutcome.Uncertain.class, attempt.outcome());
    }

    assertTrue(Files.notExists(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void reportsUncertainWhenTheMoverChangesTheDestinationThenSignalsFailure() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = Files.write(root.resolve("source.xlsx"), new byte[] {1, 2, 3, 4});

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("source.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      (source, target, disposition) -> {
                        Files.move(source, target, StandardCopyOption.REPLACE_EXISTING);
                        throw new IOException("synthetic post-move failure");
                      })
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      verifiedStage()));

      PublicationOutcome.Uncertain outcome =
          assertInstanceOf(PublicationOutcome.Uncertain.class, attempt.outcome());
      assertEquals(
          PublicationOutcome.Recovery.INSPECT_DESTINATION_DO_NOT_RETRY, outcome.recovery());
      assertEquals(
          GridGrindProblemCode.PUBLICATION_UNCERTAIN,
          GridGrindProblemCodeClassifier.codeFor(
              new WorkbookPublicationException(outcome, attempt.failure())));
    }

    assertArrayEquals(new byte[] {6, 7, 8}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void recoversAPublishedOutcomeWhenTheSiblingAndFinalDestinationBothVerifyAfterMoverFailure()
      throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      PublishedPublicationAttempt attempt =
          assertInstanceOf(
              PublishedPublicationAttempt.class,
              new RequestPathPublication(
                      (source, target, disposition) -> {
                        Files.copy(source, target);
                        throw new IOException("mover failed after destination copy");
                      })
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      assertEquals(destination.toString(), attempt.outcome().executionPath());
    }

    assertArrayEquals(new byte[] {6, 7, 8}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void reportsUncertainWhenPostMoveVerificationFindsDifferentFinalBytes() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = Files.write(root.resolve("source.xlsx"), new byte[] {1, 2, 3, 4});

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("source.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      (source, target, disposition) -> {
                        moveAtomically(source, target, disposition);
                        Files.write(target, new byte[] {9});
                      })
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      verifiedStage()));

      PublicationOutcome.Uncertain outcome =
          assertInstanceOf(PublicationOutcome.Uncertain.class, attempt.outcome());
      assertEquals(destination.toString(), outcome.executionPath());
    }

    assertArrayEquals(new byte[] {9}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void publicationFailureAndCreateNewConflictExceptionsRetainOnlyTruthfulOutcomeStates() {
    PublicationOutcome.NotPublished notPublished =
        new PublicationOutcome.NotPublished(PublicationOutcome.DestinationState.PRESERVED);
    IOException cause = new IOException("publication failed");
    WorkbookPublicationException failure = new WorkbookPublicationException(notPublished, cause);

    assertEquals(notPublished, failure.publication());
    assertEquals(GridGrindProblemCode.IO_ERROR, GridGrindProblemCodeClassifier.codeFor(failure));
    assertEquals(
        GridGrindProblemCode.PUBLICATION_UNCERTAIN,
        GridGrindProblemCodeClassifier.codeFor(
            new WorkbookPublicationException(
                new PublicationOutcome.Uncertain(
                    "output.xlsx", PublicationOutcome.Recovery.INSPECT_DESTINATION_DO_NOT_RETRY),
                cause)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookPublicationException(new PublicationOutcome.NotAttempted(), cause));
    assertEquals(
        "Workbook output path already exists and ifExists=REJECT: output.xlsx",
        new OutputPathAlreadyExistsException("output.xlsx").getMessage());
    assertEquals(cause, new OutputPathAlreadyExistsException("output.xlsx", cause).getCause());
  }

  @Test
  void retainsUncertainFailuresAndSuppressesPrivateSiblingCleanupNoise() {
    PublicationOutcome.Uncertain uncertain =
        new PublicationOutcome.Uncertain(
            root.resolve("output.xlsx").toString(),
            PublicationOutcome.Recovery.INSPECT_DESTINATION_DO_NOT_RETRY);
    IOException failure = new IOException("publication state unavailable");

    assertEquals(uncertain, new FailedPublicationAttempt(uncertain, failure).outcome());
    assertThrows(
        IllegalArgumentException.class,
        () -> new FailedPublicationAttempt(new PublicationOutcome.NotAttempted(), failure));

    AtomicInteger cleanupAttempts = new AtomicInteger();
    RequestPathPublication.deleteSiblingQuietly(
        Path.of(".publication-sibling.xlsx"),
        sibling -> {
          cleanupAttempts.incrementAndGet();
          throw new java.nio.file.NoSuchFileException(sibling.toString());
        });
    RequestPathPublication.deleteSiblingQuietly(
        Path.of(".publication-sibling.xlsx"),
        sibling -> {
          cleanupAttempts.incrementAndGet();
          throw new IOException("synthetic private cleanup failure");
        });

    assertEquals(2, cleanupAttempts.get());
  }
}
