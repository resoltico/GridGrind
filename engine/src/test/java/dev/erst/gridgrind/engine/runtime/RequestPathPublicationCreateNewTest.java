package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.assertNoPrivatePublicationSibling;
import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.copySiblingExactly;
import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.verifiedStage;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests truthful outcomes for create-new publication attempts. */
class RequestPathPublicationCreateNewTest {
  @TempDir Path root;

  @Test
  void publishesVerifiedCreateNewArtifactWithoutLeavingThePrivateSibling() throws Exception {
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
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      assertEquals(destination.toString(), attempt.outcome().executionPath());
      assertEquals(3, attempt.outcome().byteSize());
    }

    assertArrayEquals(new byte[] {6, 7, 8}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void reportsUncertainWhenPrimitiveInvocationThrowsWithoutProvenOccurrence() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
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
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      var outcome = assertInstanceOf(PublicationOutcome.Uncertain.class, attempt.outcome());
      assertEquals(
          PublicationOutcome.CommitObservation.NOT_ESTABLISHED, outcome.commitObservation());
    }

    assertTrue(Files.notExists(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void doesNotAttemptMutationWhenBoundDestinationAlreadyChanged() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      Files.write(destination, new byte[] {9});

      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication()
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      assertInstanceOf(PublicationOutcome.NotAttempted.class, attempt.outcome());
    }

    assertArrayEquals(new byte[] {9}, Files.readAllBytes(destination));
  }

  @Test
  void doesNotAttemptMutationWhenDestinationChangesDuringPrivateStaging() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      RequestPathPublicationTestSupport::commitAtomically,
                      (stagedArtifact, targetBinding, siblingName) -> {
                        copySiblingExactly(stagedArtifact, targetBinding, siblingName);
                        Files.write(targetBinding.resolvedPath(), new byte[] {9});
                      })
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      assertInstanceOf(PublicationOutcome.NotAttempted.class, attempt.outcome());
    }

    assertArrayEquals(new byte[] {9}, Files.readAllBytes(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void leavesCreateNewDestinationAbsentWhenPrivateSiblingCopyFailsAfterOneByte() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      RequestPathPublicationTestSupport::commitAtomically,
                      RequestPathPublicationTestSupport::writeOneByteThenFail)
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      assertInstanceOf(PublicationOutcome.NotAttempted.class, attempt.outcome());
    }

    assertTrue(Files.notExists(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void leavesCreateNewDestinationAbsentWhenPrivateSiblingDigestDoesNotMatch() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      RequestPathPublicationTestSupport::commitAtomically,
                      RequestPathPublicationTestSupport::writeDigestMismatchSibling)
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      assertInstanceOf(PublicationOutcome.NotAttempted.class, attempt.outcome());
    }

    assertTrue(Files.notExists(destination));
    assertNoPrivatePublicationSibling(root);
  }
}
