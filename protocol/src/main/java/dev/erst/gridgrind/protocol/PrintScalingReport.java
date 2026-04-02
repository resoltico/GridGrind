package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Print scaling state captured from one sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PrintScalingReport.Automatic.class, name = "AUTOMATIC"),
  @JsonSubTypes.Type(value = PrintScalingReport.Fit.class, name = "FIT")
})
public sealed interface PrintScalingReport
    permits PrintScalingReport.Automatic, PrintScalingReport.Fit {
  /** Sheet uses Excel's default scaling instead of fit-to-page counts. */
  record Automatic() implements PrintScalingReport {}

  /** Sheet fits printed content into the provided page counts. */
  record Fit(int widthPages, int heightPages) implements PrintScalingReport {}
}
