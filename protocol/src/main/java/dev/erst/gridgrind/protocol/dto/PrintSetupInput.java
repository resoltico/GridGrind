package dev.erst.gridgrind.protocol.dto;

import java.util.List;
import java.util.Objects;

/** Advanced page-setup payload nested under print-layout authoring. */
public record PrintSetupInput(
    PrintMarginsInput margins,
    Boolean horizontallyCentered,
    Boolean verticallyCentered,
    Integer paperSize,
    Boolean draft,
    Boolean blackAndWhite,
    Integer copies,
    Boolean useFirstPageNumber,
    Integer firstPageNumber,
    List<Integer> rowBreaks,
    List<Integer> columnBreaks) {
  /** Returns the default advanced page-setup payload for an unconfigured sheet. */
  public static PrintSetupInput defaults() {
    return new PrintSetupInput(
        new PrintMarginsInput(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d),
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

  public PrintSetupInput {
    margins = margins == null ? defaults().margins() : margins;
    horizontallyCentered = Boolean.TRUE.equals(horizontallyCentered);
    verticallyCentered = Boolean.TRUE.equals(verticallyCentered);
    paperSize = paperSize == null ? 0 : paperSize;
    draft = Boolean.TRUE.equals(draft);
    blackAndWhite = Boolean.TRUE.equals(blackAndWhite);
    copies = copies == null ? 0 : copies;
    useFirstPageNumber = Boolean.TRUE.equals(useFirstPageNumber);
    firstPageNumber = firstPageNumber == null ? 0 : firstPageNumber;
    rowBreaks = copyIndexes(rowBreaks == null ? List.of() : rowBreaks, "rowBreaks");
    columnBreaks = copyIndexes(columnBreaks == null ? List.of() : columnBreaks, "columnBreaks");
    Objects.requireNonNull(margins, "margins must not be null");
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
      Objects.requireNonNull(value, fieldName + " must not contain null values");
      if (value < 0) {
        throw new IllegalArgumentException(fieldName + " must not contain negative indexes");
      }
    }
    return copy;
  }
}
