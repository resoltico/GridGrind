package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** One requested print scaling state in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintScalingInput.Automatic.class, name = "AUTOMATIC"),
  @JsonSubTypes.Type(value = PrintScalingInput.Fit.class, name = "FIT")
})
public sealed interface PrintScalingInput
    permits PrintScalingInput.Automatic, PrintScalingInput.Fit {
  /** Sheet uses Excel's default scaling instead of fit-to-page counts. */
  record Automatic() implements PrintScalingInput {}

  /**
   * Sheet fits printed content into the provided page counts.
   *
   * <p>A value of {@code 0} on one axis keeps that axis unconstrained, matching Excel's fit
   * semantics.
   */
  record Fit(int widthPages, int heightPages) implements PrintScalingInput {
    public Fit {
      requirePageCount(widthPages, "widthPages");
      requirePageCount(heightPages, "heightPages");
    }
  }

  private static void requirePageCount(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
    if (value > Short.MAX_VALUE) {
      throw new IllegalArgumentException(
          fieldName + " must not exceed " + Short.MAX_VALUE + ": " + value);
    }
  }
}
