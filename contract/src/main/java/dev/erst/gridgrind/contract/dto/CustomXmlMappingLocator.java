package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Identifies one workbook custom-XML mapping by stable map id, authored name, or both. */
public record CustomXmlMappingLocator(Long mapId, String name) {
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
