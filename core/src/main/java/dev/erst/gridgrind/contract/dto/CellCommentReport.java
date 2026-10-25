package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Comment metadata associated with one concrete cell address. */
public record CellCommentReport(String address, CommentReport comment) {
  public CellCommentReport {
    Objects.requireNonNull(address, "address must not be null");
    Objects.requireNonNull(comment, "comment must not be null");
    if (address.isBlank()) {
      throw new IllegalArgumentException("address must not be blank");
    }
  }
}
