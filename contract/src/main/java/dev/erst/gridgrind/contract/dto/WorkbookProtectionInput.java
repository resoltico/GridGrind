package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Workbook-protection payload covering workbook and revisions lock state plus passwords. */
public record WorkbookProtectionInput(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> workbookPassword,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> revisionsPassword) {
  public WorkbookProtectionInput {
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
