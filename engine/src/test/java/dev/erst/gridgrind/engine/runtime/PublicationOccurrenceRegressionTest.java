package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Distinguishes this run's commit occurrence from content equality and later observations. */
class PublicationOccurrenceRegressionTest {
  @TempDir Path root;

  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  void competingLeafAfterTheLastCheckCannotBeOverwrittenEvenWhenIdentical(boolean identical)
      throws Exception {
    Path staged = Files.writeString(root.resolve("staged.xlsx"), "intended");
    Path destination = root.resolve("output.xlsx");
    String competitor = identical ? "intended" : "concurrent writer";
    try (var binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      var publisher =
          new RequestPathPublication((source, target, disposition) -> {
            Files.writeString(target, competitor, StandardOpenOption.CREATE_NEW);
            RequestPathPublication.commitAtomically(source, target, disposition);
          });
      var failed =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              publisher.publish(
                  binding,
                  staged,
                  WorkbookArtifactWriteDisposition.CREATE_NEW,
                  new PublicationOutcome.StagedArtifactVerification()));
      assertEquals(
          new PublicationOutcome.NotPublished(
              PublicationOutcome.DestinationState.COLLISION_OBSERVED),
          failed.outcome());
      assertEquals(competitor, Files.readString(destination));
    }
  }

  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  void preexistingLeafRejectsRegardlessOfDigestEquality(boolean identical) throws Exception {
    Path staged = Files.writeString(root.resolve("staged.xlsx"), "intended");
    String original = identical ? "intended" : "original writer";
    Path destination = Files.writeString(root.resolve("output.xlsx"), original);
    try (var binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      var failed =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              new RequestPathPublication()
                  .publish(
                      binding,
                      staged,
                      WorkbookArtifactWriteDisposition.CREATE_NEW,
                      new PublicationOutcome.StagedArtifactVerification()));
      assertEquals(
          new PublicationOutcome.NotPublished(
              PublicationOutcome.DestinationState.COLLISION_OBSERVED),
          failed.outcome());
      assertEquals(original, Files.readString(destination));
    }
  }

  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  void successfulPrimitiveReturnSurvivesMissingOrChangedReadback(boolean missing) throws Exception {
    Path staged = Files.writeString(root.resolve("staged.xlsx"), "intended");
    try (var binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      var publisher =
          new RequestPathPublication((source, target, disposition) -> {
            RequestPathPublication.commitAtomically(source, target, disposition);
            if (missing) Files.delete(target);
            else Files.writeString(target, "changed after own commit");
          });
      var failed =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              publisher.publish(
                  binding,
                  staged,
                  WorkbookArtifactWriteDisposition.CREATE_NEW,
                  new PublicationOutcome.StagedArtifactVerification()));
      var uncertain = assertInstanceOf(PublicationOutcome.Uncertain.class, failed.outcome());
      assertEquals(
          PublicationOutcome.CommitObservation.RETURNED_SUCCESSFULLY,
          uncertain.commitObservation());
      assertEquals(
          PublicationOutcome.Recovery.INSPECT_DESTINATION_DO_NOT_RETRY, uncertain.recovery());
    }
  }

  @Test
  void fileSynchronizationFailureRetainsSuccessfulCommitOccurrence() throws Exception {
    Path staged = Files.writeString(root.resolve("stage.xlsx"), "complete");
    try (var binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      var publisher =
          new RequestPathPublication(
              RequestPathPublication::commitAtomically,
              RequestPathPublicationTestSupport::copySiblingExactly,
              channel -> {
                throw new IOException("file force failed");
              });
      var attempt =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              publisher.publish(
                  binding,
                  staged,
                  WorkbookArtifactWriteDisposition.CREATE_NEW,
                  new PublicationOutcome.StagedArtifactVerification()));
      var outcome = assertInstanceOf(PublicationOutcome.Uncertain.class, attempt.outcome());
      assertEquals(
          PublicationOutcome.CommitObservation.RETURNED_SUCCESSFULLY, outcome.commitObservation());
      assertEquals("complete", Files.readString(root.resolve("output.xlsx")));
    }
  }

  @Test
  void ambiguousThrowNeverEstablishesOwnCommitFromMatchingBytes() throws Exception {
    Path staged = Files.writeString(root.resolve("staged.xlsx"), "intended");
    try (var binding = RequestPathBinding.bindWriteTarget("output.xlsx", root)) {
      var publisher =
          new RequestPathPublication((source, target, disposition) -> {
            Files.writeString(target, "intended", StandardOpenOption.CREATE_NEW);
            throw new IOException("synthetic ambiguous primitive failure");
          });
      var failed =
          assertInstanceOf(
              FailedPublicationAttempt.class,
              publisher.publish(
                  binding,
                  staged,
                  WorkbookArtifactWriteDisposition.CREATE_NEW,
                  new PublicationOutcome.StagedArtifactVerification()));
      var uncertain = assertInstanceOf(PublicationOutcome.Uncertain.class, failed.outcome());
      assertEquals(
          PublicationOutcome.CommitObservation.NOT_ESTABLISHED, uncertain.commitObservation());
    }
  }
}
