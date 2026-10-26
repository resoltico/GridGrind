package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;

/** Identifies one workbook custom-XML mapping by stable map id, authored name, or both. */
@CatalogType(
    id = "CustomXmlMappingLocator",
    summary =
        "One locator for an existing workbook custom-XML mapping. Supply mapId, name, or both; the locator must resolve to exactly one existing mapping.",
    displayOrder = 31)
public record CustomXmlMappingLocator(Optional<Long> mapId, Optional<String> name) {
  public CustomXmlMappingLocator {
    Objects.requireNonNull(mapId, "mapId must not be null");
    Objects.requireNonNull(name, "name must not be null");
    if (mapId.isEmpty() && name.isEmpty()) {
      throw new IllegalArgumentException("mapId or name must be provided");
    }
    if (mapId.isPresent() && mapId.orElseThrow() <= 0L) {
      throw new IllegalArgumentException("mapId must be greater than 0");
    }
    name = name.map(value -> requireNonBlank(value, "name"));
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
