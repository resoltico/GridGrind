package dev.erst.gridgrind.contract.query;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;

/** Derived factual surface-summary inspection queries. */
public sealed interface InspectionSurfaceQuery extends InspectionQuery.Surface
    permits InspectionSurfaceQuery.GetFormulaSurface,
        InspectionSurfaceQuery.GetSheetSchema,
        InspectionSurfaceQuery.GetNamedRangeSurface {

  @ProtocolTypeMetadata(
      id = "GET_FORMULA_SURFACE",
      summary = GridGrindContractText.FORMULA_SURFACE_READ_SUMMARY,
      targetSelectors = {SheetSelector.class})
  record GetFormulaSurface() implements InspectionSurfaceQuery {}

  @ProtocolTypeMetadata(
      id = "GET_SHEET_SCHEMA",
      summary = "Infer a simple schema from a rectangular sheet window.",
      targetSelectors = {RangeSelector.RectangularWindow.class})
  record GetSheetSchema() implements InspectionSurfaceQuery {}

  @ProtocolTypeMetadata(
      id = "GET_NAMED_RANGE_SURFACE",
      summary = GridGrindContractText.NAMED_RANGE_SURFACE_READ_SUMMARY,
      targetSelectors = {NamedRangeSelector.class})
  record GetNamedRangeSurface() implements InspectionSurfaceQuery {}
}
