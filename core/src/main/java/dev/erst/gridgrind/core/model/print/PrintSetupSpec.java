package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.List;
import java.util.Objects;

/** Advanced page-setup payload nested under print-layout authoring. */
@CatalogType(
    id = "PrintSetupSpec",
    summary =
        "Advanced page-setup payload nested under print-layout authoring. Supply the full authored setup block explicitly.",
    displayOrder = 65)
public record PrintSetupSpec(
    PrintMarginsSpec margins,
    boolean printGridlines,
    boolean horizontallyCentered,
    boolean verticallyCentered,
    int paperSize,
    boolean draft,
    boolean blackAndWhite,
    int copies,
    boolean useFirstPageNumber,
    int firstPageNumber,
    List<Integer> rowBreaks,
    List<Integer> columnBreaks) {
  /** Returns the default advanced page-setup payload for an unconfigured sheet. */
  public static PrintSetupSpec defaults() {
    return new PrintSetupSpec(
        new PrintMarginsSpec(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
        false,
        false,
        false,
        1,
        false,
        false,
        1,
        false,
        1,
        List.of(),
        List.of());
  }

  public PrintSetupSpec {
    Objects.requireNonNull(margins, "margins must not be null");
    rowBreaks = copyIndexes(rowBreaks, "rowBreaks");
    columnBreaks = copyIndexes(columnBreaks, "columnBreaks");
    if (paperSize < 0) {
      throw new IllegalArgumentException("paperSize must not be negative");
    }
    if (copies < 0) {
      throw new IllegalArgumentException("copies must not be negative");
    }
    if (firstPageNumber < 0) {
      throw new IllegalArgumentException("firstPageNumber must not be negative");
    }
  }

  private static List<Integer> copyIndexes(List<Integer> values, String fieldName) {
    List<Integer> copy =
        List.copyOf(Objects.requireNonNull(values, fieldName + " must not be null"));
    for (Integer value : copy) {
      if (value < 0) {
        throw new IllegalArgumentException(fieldName + " must not contain negative indexes");
      }
    }
    return copy;
  }
}
