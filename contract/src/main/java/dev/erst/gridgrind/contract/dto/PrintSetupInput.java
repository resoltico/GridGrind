package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import java.util.List;
import java.util.Objects;

/** Advanced page-setup payload nested under print-layout authoring. */
public record PrintSetupInput(
    PrintMarginsInput margins,
    Boolean printGridlines,
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
    Objects.requireNonNull(margins, "margins must not be null");
    Objects.requireNonNull(printGridlines, "printGridlines must not be null");
    Objects.requireNonNull(horizontallyCentered, "horizontallyCentered must not be null");
    Objects.requireNonNull(verticallyCentered, "verticallyCentered must not be null");
    Objects.requireNonNull(paperSize, "paperSize must not be null");
    Objects.requireNonNull(draft, "draft must not be null");
    Objects.requireNonNull(blackAndWhite, "blackAndWhite must not be null");
    Objects.requireNonNull(copies, "copies must not be null");
    Objects.requireNonNull(useFirstPageNumber, "useFirstPageNumber must not be null");
    Objects.requireNonNull(firstPageNumber, "firstPageNumber must not be null");
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

  @JsonCreator(mode = JsonCreator.Mode.DELEGATING)
  static PrintSetupInput create(PrintSetupJson json) {
    PrintSetupInput defaults = defaults();
    return new PrintSetupInput(
        json.margins() == null ? defaults.margins() : json.margins(),
        Boolean.TRUE.equals(json.printGridlines()),
        Boolean.TRUE.equals(json.horizontallyCentered()),
        Boolean.TRUE.equals(json.verticallyCentered()),
        json.paperSize() == null ? defaults.paperSize() : json.paperSize(),
        Boolean.TRUE.equals(json.draft()),
        Boolean.TRUE.equals(json.blackAndWhite()),
        json.copies() == null ? defaults.copies() : json.copies(),
        Boolean.TRUE.equals(json.useFirstPageNumber()),
        json.firstPageNumber() == null ? defaults.firstPageNumber() : json.firstPageNumber(),
        json.rowBreaks() == null ? List.of() : json.rowBreaks(),
        json.columnBreaks() == null ? List.of() : json.columnBreaks());
  }

  private record PrintSetupJson(
      PrintMarginsInput margins,
      Boolean printGridlines,
      Boolean horizontallyCentered,
      Boolean verticallyCentered,
      Integer paperSize,
      Boolean draft,
      Boolean blackAndWhite,
      Integer copies,
      Boolean useFirstPageNumber,
      Integer firstPageNumber,
      List<Integer> rowBreaks,
      List<Integer> columnBreaks) {}
}
