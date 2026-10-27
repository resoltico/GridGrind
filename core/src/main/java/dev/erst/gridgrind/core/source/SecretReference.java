package dev.erst.gridgrind.core.source;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.FieldRule;
import java.util.Objects;

/** Opaque host-resolved identifier for one secret required by an authored workbook plan. */
@CatalogType(
    id = "SecretReference",
    summary = "An opaque identifier resolved by the host without carrying secret material.")
public record SecretReference(
    @FieldRule(kind = FieldRule.Kind.STRING_PATTERN, text = SecretReference.ID_PATTERN) String id) {
  private static final String ID_PATTERN = "[A-Za-z][A-Za-z0-9._-]*";

  public SecretReference {
    Objects.requireNonNull(id, "id must not be null");
    if (!id.matches(ID_PATTERN)) {
      throw new IllegalArgumentException(
          "id must start with a letter and contain only letters, digits, '.', '_', or '-'");
    }
  }
}
