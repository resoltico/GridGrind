package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Count of one observed cell value type inside a schema column. */
public record TypeCountReport(String type, int count) {
  public TypeCountReport {
    Objects.requireNonNull(type, "type must not be null");
    if (type.isBlank()) {
      throw new IllegalArgumentException("type must not be blank");
    }
    if (count <= 0) {
      throw new IllegalArgumentException("count must be greater than 0");
    }
  }
}
