package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome;
import java.io.IOException;

/** Accumulates this run's publication occurrence before verification, callbacks, or cleanup. */
final class PublicationFacts {
  boolean stagedVerified;
  private PublicationOutcome outcome = new PublicationOutcome.NotAttempted();
  private @org.jspecify.annotations.Nullable IOException cleanupFailure;

  void attempting(String path) {
    outcome = uncertain(path, PublicationOutcome.CommitObservation.NOT_ESTABLISHED);
  }

  void returnedSuccessfully(String path) {
    outcome = uncertain(path, PublicationOutcome.CommitObservation.RETURNED_SUCCESSFULLY);
  }

  void established(PublicationOutcome value) {
    outcome = value;
  }

  PublicationOutcome outcome() {
    return outcome;
  }

  @SuppressWarnings({
    "ReferenceEquality",
    "PMD.CompareObjectsWithEquals"
  }) // Throwable self-suppression is forbidden by identity.
  void cleanupFailed(IOException failure) {
    if (cleanupFailure == null) cleanupFailure = failure;
    else if (cleanupFailure != failure && cleanupFailure.getSuppressed().length < 8)
      cleanupFailure.addSuppressed(failure);
  }

  @org.jspecify.annotations.Nullable
  IOException cleanupFailure() {
    return cleanupFailure;
  }

  private static PublicationOutcome.Uncertain uncertain(
      String path, PublicationOutcome.CommitObservation observation) {
    return new PublicationOutcome.Uncertain(
        path, PublicationOutcome.Recovery.INSPECT_DESTINATION_DO_NOT_RETRY, observation);
  }
}
