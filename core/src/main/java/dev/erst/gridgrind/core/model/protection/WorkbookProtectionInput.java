package dev.erst.gridgrind.core.model.protection;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.source.SecretReference;
import java.util.Objects;
import java.util.Optional;

/** Workbook-protection payload covering workbook and revisions lock state plus passwords. */
@CatalogType(
    id = "WorkbookProtectionInput",
    summary =
        "Workbook-protection payload covering workbook and revisions lock state plus optional passwords.",
    displayOrder = 75)
public record WorkbookProtectionInput(
    @ProtocolField(optional = true) @WireDefault("false") boolean structureLocked,
    @ProtocolField(optional = true) @WireDefault("false") boolean windowsLocked,
    @ProtocolField(optional = true) @WireDefault("false") boolean revisionsLocked,
    Optional<SecretReference> workbookPasswordRef,
    Optional<SecretReference> revisionsPasswordRef) {

  public WorkbookProtectionInput {
    Objects.requireNonNull(workbookPasswordRef, "workbookPasswordRef must not be null");
    Objects.requireNonNull(revisionsPasswordRef, "revisionsPasswordRef must not be null");
  }
}
