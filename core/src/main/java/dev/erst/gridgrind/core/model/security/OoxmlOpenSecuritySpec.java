package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.SecretReference;
import java.util.Objects;
import java.util.Optional;

/** Optional OOXML package-open settings for encrypted existing workbook sources. */
@CatalogType(
    id = "OoxmlOpenSecuritySpec",
    summary =
        "Optional OOXML package-open settings for encrypted existing workbook sources. password unlocks the encrypted OOXML package before GridGrind opens the inner .xlsx workbook.",
    displayOrder = 19)
public record OoxmlOpenSecuritySpec(Optional<SecretReference> passwordRef) {
  public OoxmlOpenSecuritySpec {
    Objects.requireNonNull(passwordRef, "passwordRef must not be null");
  }
}
