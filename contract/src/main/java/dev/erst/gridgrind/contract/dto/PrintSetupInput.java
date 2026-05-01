package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.List;
import java.util.Objects;

/** Advanced page-setup payload nested under print-layout authoring. */
public record PrintSetupInput(
    PrintMarginsInput margins,
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

  /** Creates a page-setup payload from the full authored wire shape. */
  @JsonCreator
  @SuppressWarnings("PMD.ExcessiveParameterList")
  public PrintSetupInput(
      @JsonProperty("margins") PrintMarginsInput margins,
      @JsonProperty("printGridlines") Boolean printGridlines,
      @JsonProperty("horizontallyCentered") Boolean horizontallyCentered,
      @JsonProperty("verticallyCentered") Boolean verticallyCentered,
      @JsonProperty("paperSize") Integer paperSize,
      @JsonProperty("draft") Boolean draft,
      @JsonProperty("blackAndWhite") Boolean blackAndWhite,
      @JsonProperty("copies") Integer copies,
      @JsonProperty("useFirstPageNumber") Boolean useFirstPageNumber,
      @JsonProperty("firstPageNumber") Integer firstPageNumber,
      @JsonProperty("rowBreaks") List<Integer> rowBreaks,
      @JsonProperty("columnBreaks") List<Integer> columnBreaks) {
    this(
        Objects.requireNonNull(margins, "margins must not be null"),
        Objects.requireNonNull(printGridlines, "printGridlines must not be null").booleanValue(),
        Objects.requireNonNull(horizontallyCentered, "horizontallyCentered must not be null")
            .booleanValue(),
        Objects.requireNonNull(verticallyCentered, "verticallyCentered must not be null")
            .booleanValue(),
        Objects.requireNonNull(paperSize, "paperSize must not be null").intValue(),
        Objects.requireNonNull(draft, "draft must not be null").booleanValue(),
        Objects.requireNonNull(blackAndWhite, "blackAndWhite must not be null").booleanValue(),
        Objects.requireNonNull(copies, "copies must not be null").intValue(),
        Objects.requireNonNull(useFirstPageNumber, "useFirstPageNumber must not be null")
            .booleanValue(),
        Objects.requireNonNull(firstPageNumber, "firstPageNumber must not be null").intValue(),
        Objects.requireNonNull(rowBreaks, "rowBreaks must not be null"),
        Objects.requireNonNull(columnBreaks, "columnBreaks must not be null"));
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
