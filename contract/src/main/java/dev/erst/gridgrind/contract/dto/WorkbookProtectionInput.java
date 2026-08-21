package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;
import java.util.Optional;

/** Workbook-protection payload covering workbook and revisions lock state plus passwords. */
public record WorkbookProtectionInput(
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean structureLocked,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean windowsLocked,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean revisionsLocked,
    @ProtocolField(optional = true, secret = true) @JsonInclude(JsonInclude.Include.NON_ABSENT)
        Optional<String> workbookPassword,
    @ProtocolField(optional = true, secret = true) @JsonInclude(JsonInclude.Include.NON_ABSENT)
        Optional<String> revisionsPassword) {
  /** Reads workbook protection while defaulting omitted booleans to false and passwords empty. */
  @JsonCreator
  public WorkbookProtectionInput(
      @JsonProperty("structureLocked") Boolean structureLocked,
      @JsonProperty("windowsLocked") Boolean windowsLocked,
      @JsonProperty("revisionsLocked") Boolean revisionsLocked,
      @JsonProperty("workbookPassword") Optional<String> workbookPassword,
      @JsonProperty("revisionsPassword") Optional<String> revisionsPassword) {
    this(
        ProtocolBooleanDefault.FALSE.resolve(structureLocked),
        ProtocolBooleanDefault.FALSE.resolve(windowsLocked),
        ProtocolBooleanDefault.FALSE.resolve(revisionsLocked),
        emptyIfNull(workbookPassword),
        emptyIfNull(revisionsPassword));
  }

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

  private static Optional<String> emptyIfNull(Optional<String> value) {
    return value == null ? Optional.empty() : value;
  }
}
