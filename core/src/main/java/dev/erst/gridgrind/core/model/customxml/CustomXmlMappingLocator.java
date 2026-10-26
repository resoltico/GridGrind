package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import java.util.Objects;

/** Identifies one workbook custom-XML mapping by stable map id, authored name, or both. */
@CatalogType(
    id = "CustomXmlMappingLocator",
    summary =
        "One locator for an existing workbook custom-XML mapping. Supply mapId, name, or both; the locator must resolve to exactly one existing mapping.",
    displayOrder = 31)
public record CustomXmlMappingLocator(
    @ProtocolField(optional = true) Long mapId, @ProtocolField(optional = true) String name) {
  public CustomXmlMappingLocator {
    if (mapId == null && name == null) {
      throw new IllegalArgumentException("mapId or name must be provided");
    }
    if (mapId != null && mapId <= 0L) {
      throw new IllegalArgumentException("mapId must be greater than 0");
    }
    if (name != null) {
      name = requireNonBlank(name, "name");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
