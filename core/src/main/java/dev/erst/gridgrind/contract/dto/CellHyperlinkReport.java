package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Hyperlink metadata associated with one concrete cell address. */
public record CellHyperlinkReport(String address, HyperlinkTarget hyperlink) {
  public CellHyperlinkReport {
    Objects.requireNonNull(address, "address must not be null");
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");
    if (address.isBlank()) {
      throw new IllegalArgumentException("address must not be blank");
    }
  }
}
