package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Optional interpretation of a numeric fact; the original number belongs to the VALUE facet. */
@Union(discriminator = "type", catalog = true, displayOrder = 85)
public sealed interface CellTemporalReport
    permits CellTemporalReport.NotTemporal,
        CellTemporalReport.Interpreted,
        CellTemporalReport.Unrepresentable {
  @Variant(id = "NOT_TEMPORAL", summary = "The observed format does not denote a temporal value.")
  record NotTemporal() implements CellTemporalReport {}

  @Variant(
      id = "INTERPRETED",
      summary = "The numeric fact has a supported civil DATE, TIME or DATE_TIME interpretation.")
  record Interpreted(CellTemporalKind kind, String isoValue) implements CellTemporalReport {
    public Interpreted {
      Objects.requireNonNull(kind);
      Objects.requireNonNull(isoValue);
      try {
        switch (kind) {
          case DATE -> {
            if (!TemporalSemantics.supportedDateShape(java.time.LocalDate.parse(isoValue)))
              throw new IllegalArgumentException("Invalid DATE interpretation.");
          }
          case DATE_TIME -> {
            var value = java.time.LocalDateTime.parse(isoValue);
            if (!TemporalSemantics.supportedDateShape(value.toLocalDate())
                || !TemporalSemantics.millisecondTime(value.toLocalTime()))
              throw new IllegalArgumentException("Invalid DATE_TIME interpretation.");
          }
          case TIME -> {
            if (!TemporalSemantics.millisecondTime(java.time.LocalTime.parse(isoValue)))
              throw new IllegalArgumentException(
                  "TIME interpretation supports millisecond precision.");
          }
          case ELAPSED ->
              throw new IllegalArgumentException(
                  "Elapsed formats do not have a supported civil interpretation.");
        }
      } catch (java.time.DateTimeException invalid) {
        throw new IllegalArgumentException(
            "isoValue must be a civil ISO value matching its temporal kind.", invalid);
      }
    }
  }

  @Variant(
      id = "UNREPRESENTABLE",
      summary =
          "The format is temporal, but its numeric fact has no supported interpretation; request"
              + " VALUE for the original number.")
  record Unrepresentable(CellTemporalKind kind, CellTemporalReason reason)
      implements CellTemporalReport {
    public Unrepresentable {
      Objects.requireNonNull(kind);
      Objects.requireNonNull(reason);
      if ((kind == CellTemporalKind.ELAPSED)
          != (reason == CellTemporalReason.UNSUPPORTED_ELAPSED_FORMAT))
        throw new IllegalArgumentException("Elapsed format kind and refusal reason must agree.");
    }
  }
}
