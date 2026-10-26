package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** One requested print scaling state in protocol form. */
@Union(discriminator = "type", catalog = true, displayOrder = 67)
public sealed interface PrintScalingSpec
    permits dev.erst.gridgrind.core.model.print.PrintScalingSpec.Automatic,
        dev.erst.gridgrind.core.model.print.PrintScalingSpec.Fit {
  /** Sheet uses Excel's default scaling instead of fit-to-page counts. */
  @Variant(
      id = "AUTOMATIC",
      summary = "Sheet uses Excel's default scaling instead of fit-to-page counts.")
  record Automatic() implements PrintScalingSpec {}

  /**
   * Sheet fits printed content into the provided page counts.
   *
   * <p>A value of {@code 0} on one axis keeps that axis unconstrained, matching Excel's fit
   * semantics.
   */
  @Variant(
      id = "FIT",
      summary =
          "Sheet fits printed content into the provided page counts. A value of 0 on one axis keeps that axis unconstrained.")
  record Fit(int widthPages, int heightPages) implements PrintScalingSpec {
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
