package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Optional temporal interpretation projected from one numeric readback value. */
@CatalogType(
    id = "CellTemporalReport",
    summary =
        "Projected temporal facet for one numeric cell value, including whether Excel formatting marks it as date-like and any derived DATE, TIME, or DATE_TIME semantic kind.",
    displayOrder = 85)
public record CellTemporalReport(
    boolean isDate, Optional<CellTemporalKind> kind, Optional<String> isoValue) {
  public CellTemporalReport {
    kind = Objects.requireNonNullElseGet(kind, Optional::empty);
    isoValue = Objects.requireNonNullElseGet(isoValue, Optional::empty);
    isoValue.ifPresent(
        value -> {
          TextValidation.requireNonBlank(value, "isoValue");
        });
    if (isDate) {
      if (kind.isEmpty()) {
        throw new IllegalArgumentException("kind must be present when isDate is true");
      }
      if (isoValue.isEmpty()) {
        throw new IllegalArgumentException("isoValue must be present when isDate is true");
      }
    } else if (kind.isPresent() || isoValue.isPresent()) {
      throw new IllegalArgumentException("kind and isoValue must be omitted when isDate is false");
    }
  }

  /** Returns the explicit non-temporal facet value for numeric cells that are not dates. */
  public static CellTemporalReport notDate() {
    return new CellTemporalReport(false, Optional.empty(), Optional.empty());
  }

  /** Returns the temporal facet value for a numeric cell whose format denotes a date or time. */
  public static CellTemporalReport temporal(CellTemporalKind kind, String isoValue) {
    return new CellTemporalReport(true, Optional.of(kind), Optional.of(isoValue));
  }
}
