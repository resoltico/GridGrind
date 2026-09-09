package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.assertNoPrivatePublicationSibling;
import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.copySiblingExactly;
import static dev.erst.gridgrind.engine.runtime.RequestPathPublicationTestSupport.verifiedStage;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
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
  void leavesCreateNewDestinationAbsentWhenAtomicPublicationDoesNotBegin() throws Exception {
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

      PublicationOutcome.NotPublished outcome =
          assertInstanceOf(PublicationOutcome.NotPublished.class, attempt.outcome());
      assertEquals(PublicationOutcome.DestinationState.ABSENT, outcome.destinationState());
    }

    assertTrue(Files.notExists(destination));
    assertNoPrivatePublicationSibling(root);
  }

  @Test
  void reportsUncertainWhenTheBoundDestinationChangesBeforePublicationStateIsCaptured()
      throws Exception {
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

      assertInstanceOf(PublicationOutcome.Uncertain.class, attempt.outcome());
    }

    assertArrayEquals(new byte[] {9}, Files.readAllBytes(destination));
  }

  @Test
  void reportsUncertainWhenADestinationAppearsAfterCreateNewStateIsCaptured() throws Exception {
    Path staged = Files.write(root.resolve("staged.xlsx"), new byte[] {6, 7, 8});
    Path destination = root.resolve("output.xlsx");

    try (RequestPathBinding binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      FailedPublicationAttempt attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication(
                      RequestPathPublicationTestSupport::moveAtomically,
                      (stagedArtifact, targetBinding, siblingName) -> {
                        copySiblingExactly(stagedArtifact, targetBinding, siblingName);
                        Files.write(targetBinding.resolvedPath(), new byte[] {9});
                      })
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      assertInstanceOf(PublicationOutcome.Uncertain.class, attempt.outcome());
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
                      RequestPathPublicationTestSupport::moveAtomically,
                      RequestPathPublicationTestSupport::writeOneByteThenFail)
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      PublicationOutcome.NotPublished outcome =
          assertInstanceOf(PublicationOutcome.NotPublished.class, attempt.outcome());
      assertEquals(PublicationOutcome.DestinationState.ABSENT, outcome.destinationState());
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
                      RequestPathPublicationTestSupport::moveAtomically,
                      RequestPathPublicationTestSupport::writeDigestMismatchSibling)
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      verifiedStage()));

      PublicationOutcome.NotPublished outcome =
          assertInstanceOf(PublicationOutcome.NotPublished.class, attempt.outcome());
      assertEquals(PublicationOutcome.DestinationState.ABSENT, outcome.destinationState());
    }

    assertTrue(Files.notExists(destination));
    assertNoPrivatePublicationSibling(root);
  }
}
