package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Print scaling state captured from one sheet. */
@Union(discriminator = "type")
public sealed interface PrintScalingReport
    permits PrintScalingReport.Automatic, PrintScalingReport.Fit {
  /** Sheet uses Excel's default scaling instead of fit-to-page counts. */
  @Variant(
      id = "AUTOMATIC",
      summary = "Sheet uses Excel's default scaling instead of fit-to-page counts.")
  record Automatic() implements PrintScalingReport {}

  /** Sheet fits printed content into the provided page counts. */
  @Variant(id = "FIT", summary = "Sheet fits printed content into the provided page counts.")
  record Fit(int widthPages, int heightPages) implements PrintScalingReport {}
}
