package dev.erst.gridgrind.engine.cells;

import dev.erst.gridgrind.core.limits.Limit;
import java.util.Objects;

/** Centralized Excel cell-text ceiling checks for authored and reconstructed sheet content. */
@SuppressWarnings("PMD.CommentRequired")
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_TEXT_CHARACTERS_PER_CELL
})
public final class CellTextLimits {

  private CellTextLimits() {}

  public static void requireSupportedLength(String text, String fieldName) {
    Objects.requireNonNull(text, fieldName + " must not be null");
    if (text.length() > Limit.MAX_CELL_TEXT_LENGTH) {
      throw new IllegalArgumentException(
          fieldName
              + " must not exceed "
              + Limit.MAX_CELL_TEXT_LENGTH
              + " characters (Excel cell text limit)");
    }
  }
}
