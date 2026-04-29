package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** Workbook-protection payload covering workbook and revisions lock state plus passwords. */
public record WorkbookProtectionInput(
    Boolean structureLocked,
    Boolean windowsLocked,
    Boolean revisionsLocked,
    String workbookPassword,
    String revisionsPassword) {
  public WorkbookProtectionInput {
    Objects.requireNonNull(structureLocked, "structureLocked must not be null");
    Objects.requireNonNull(windowsLocked, "windowsLocked must not be null");
    Objects.requireNonNull(revisionsLocked, "revisionsLocked must not be null");
    if (workbookPassword != null && workbookPassword.isBlank()) {
      throw new IllegalArgumentException("workbookPassword must not be blank");
    }
    if (revisionsPassword != null && revisionsPassword.isBlank()) {
      throw new IllegalArgumentException("revisionsPassword must not be blank");
    }
  }

  @JsonCreator
  static WorkbookProtectionInput create(
      @JsonProperty("structureLocked") Boolean structureLocked,
      @JsonProperty("windowsLocked") Boolean windowsLocked,
      @JsonProperty("revisionsLocked") Boolean revisionsLocked,
      @JsonProperty("workbookPassword") String workbookPassword,
      @JsonProperty("revisionsPassword") String revisionsPassword) {
    return new WorkbookProtectionInput(
        Boolean.TRUE.equals(structureLocked),
        Boolean.TRUE.equals(windowsLocked),
        Boolean.TRUE.equals(revisionsLocked),
        workbookPassword,
        revisionsPassword);
  }
}
