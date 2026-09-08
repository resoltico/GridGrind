package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome;
import java.io.IOException;
import java.util.Objects;

/** Carries the truthful publication state when final artifact publication cannot complete. */
final class WorkbookPublicationException extends IOException {
  private static final long serialVersionUID = 1L;

  private final PublicationOutcome publication;

  WorkbookPublicationException(PublicationOutcome publication, IOException cause) {
    super(Objects.requireNonNull(cause, "cause must not be null").getMessage(), cause);
    if (!(publication instanceof PublicationOutcome.NotPublished)
        && !(publication instanceof PublicationOutcome.Uncertain)) {
      throw new IllegalArgumentException("publication failure must be NOT_PUBLISHED or UNCERTAIN");
    }
    this.publication = publication;
  }

  PublicationOutcome publication() {
    return publication;
  }
}
