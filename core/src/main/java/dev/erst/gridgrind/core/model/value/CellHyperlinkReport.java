package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import java.util.Objects;

/** Hyperlink metadata associated with one concrete cell address. */
public record CellHyperlinkReport(CellAddress address, HyperlinkTarget hyperlink) {
  public CellHyperlinkReport {
    Objects.requireNonNull(address, "address must not be null");
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");
  }
}
