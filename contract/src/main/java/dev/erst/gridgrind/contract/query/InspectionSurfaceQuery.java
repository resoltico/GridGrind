package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.catalog.GridGrindInspectionContractText;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import java.util.Objects;
import java.util.Optional;

/** Derived factual surface-summary inspection queries. */
public sealed interface InspectionSurfaceQuery extends InspectionQuery.Surface
    permits InspectionSurfaceQuery.GetFormulaSurface,
        InspectionSurfaceQuery.GetSheetSchema,
        InspectionSurfaceQuery.GetNamedRangeSurface {

  @ProtocolTypeMetadata(
      id = "GET_FORMULA_SURFACE",
      summary = GridGrindInspectionContractText.FORMULA_SURFACE_READ_SUMMARY,
      targetSelectors = {SheetSelector.class})
  record GetFormulaSurface() implements InspectionSurfaceQuery {}

  @ProtocolTypeMetadata(
      id = "GET_SHEET_SCHEMA",
      summary =
          "Infer a simple schema from a rectangular sheet window."
              + " Omit projection for the default compact VALUE readback.",
      optionalFields = {"projection"},
      targetSelectors = {RangeSelector.RectangularWindow.class})
  record GetSheetSchema(
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellReadProjection> projection)
      implements InspectionSurfaceQuery {
    /** Creates a schema query that relies on the default cell-read projection. */
    public GetSheetSchema() {
      this(Optional.empty());
    }

    public GetSheetSchema {
      projection = Objects.requireNonNullElseGet(projection, Optional::empty);
    }

    /** Returns the effective projection after applying the default when omitted on the wire. */
    public CellReadProjection resolvedProjection() {
      return projection.orElseGet(CellReadProjection::defaults);
    }
  }

  @ProtocolTypeMetadata(
      id = "GET_NAMED_RANGE_SURFACE",
      summary = GridGrindInspectionContractText.NAMED_RANGE_SURFACE_READ_SUMMARY,
      targetSelectors = {NamedRangeSelector.class})
  record GetNamedRangeSurface() implements InspectionSurfaceQuery {}
}
