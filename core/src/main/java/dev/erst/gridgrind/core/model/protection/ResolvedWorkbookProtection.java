package dev.erst.gridgrind.core.model.protection;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Materialized workbook/revision protection credentials and lock settings. */
public record ResolvedWorkbookProtection(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    Optional<String> workbookPassword,
    Optional<String> revisionsPassword) {
  public ResolvedWorkbookProtection {
    Objects.requireNonNull(workbookPassword, "workbookPassword must not be null");
    Objects.requireNonNull(revisionsPassword, "revisionsPassword must not be null");
    workbookPassword.ifPresent(
        value -> {
          TextValidation.requireNonBlank(value, "workbookPassword");
        });
    revisionsPassword.ifPresent(
        value -> {
          TextValidation.requireNonBlank(value, "revisionsPassword");
        });
  }

  /** Renders lock state without exposing materialized credentials. */
  @Override
  public String toString() {
    return "ResolvedWorkbookProtection[structureLocked="
        + structureLocked
        + ", windowsLocked="
        + windowsLocked
        + ", revisionsLocked="
        + revisionsLocked
        + ", workbookPassword="
        + (workbookPassword.isPresent() ? "redacted" : "absent")
        + ", revisionsPassword="
        + (revisionsPassword.isPresent() ? "redacted" : "absent")
        + "]";
  }
}
