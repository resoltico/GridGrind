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
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<SecretReference> workbookPasswordRef,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<SecretReference> revisionsPasswordRef) {
  /** Reads workbook protection while defaulting omitted booleans to false and passwords empty. */
  @JsonCreator
  public WorkbookProtectionInput(
      @JsonProperty("structureLocked") Boolean structureLocked,
      @JsonProperty("windowsLocked") Boolean windowsLocked,
      @JsonProperty("revisionsLocked") Boolean revisionsLocked,
      @JsonProperty("workbookPasswordRef") Optional<SecretReference> workbookPasswordRef,
      @JsonProperty("revisionsPasswordRef") Optional<SecretReference> revisionsPasswordRef) {
    this(
        ProtocolBooleanDefault.FALSE.resolve(structureLocked),
        ProtocolBooleanDefault.FALSE.resolve(windowsLocked),
        ProtocolBooleanDefault.FALSE.resolve(revisionsLocked),
        emptyIfNull(workbookPasswordRef),
        emptyIfNull(revisionsPasswordRef));
  }

  public WorkbookProtectionInput {
    Objects.requireNonNull(workbookPasswordRef, "workbookPasswordRef must not be null");
    Objects.requireNonNull(revisionsPasswordRef, "revisionsPasswordRef must not be null");
  }

  private static Optional<SecretReference> emptyIfNull(Optional<SecretReference> value) {
    return value == null ? Optional.empty() : value;
  }
}
