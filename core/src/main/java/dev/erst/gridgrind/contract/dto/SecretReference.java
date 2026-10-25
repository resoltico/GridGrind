package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Opaque host-resolved identifier for one secret required by an authored workbook plan. */
public record SecretReference(String id) {
  public SecretReference {
    Objects.requireNonNull(id, "id must not be null");
    if (!id.matches("[A-Za-z][A-Za-z0-9._-]*")) {
      throw new IllegalArgumentException(
          "id must start with a letter and contain only letters, digits, '.', '_', or '-'");
    }
  }
}
