package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Mutable-workbook protection payload covering workbook and revisions locks plus passwords. */
public record ExcelWorkbookProtectionSettings(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    Optional<String> workbookPassword,
    Optional<String> revisionsPassword) {
  public ExcelWorkbookProtectionSettings {
    Objects.requireNonNull(workbookPassword, "workbookPassword must not be null");
    Objects.requireNonNull(revisionsPassword, "revisionsPassword must not be null");
    workbookPassword.ifPresent(
        value -> {
          if (value.isBlank()) {
            throw new IllegalArgumentException("workbookPassword must not be blank");
          }
        });
    revisionsPassword.ifPresent(
        value -> {
          if (value.isBlank()) {
            throw new IllegalArgumentException("revisionsPassword must not be blank");
          }
        });
  }
}
