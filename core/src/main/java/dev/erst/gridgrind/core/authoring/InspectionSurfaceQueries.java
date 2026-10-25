package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.plan.InspectionSurfaceQuery;

/** Canonical inspection-surface factories kept internal to the Java authoring surface. */
final class InspectionSurfaceQueries {
  private InspectionSurfaceQueries() {}

  static InspectionSurfaceQuery.GetFormulaSurface formulaSurface() {
    return new InspectionSurfaceQuery.GetFormulaSurface();
  }

  static InspectionSurfaceQuery.GetSheetSchema sheetSchema() {
    return new InspectionSurfaceQuery.GetSheetSchema();
  }

  static InspectionSurfaceQuery.GetNamedRangeSurface namedRangeSurface() {
    return new InspectionSurfaceQuery.GetNamedRangeSurface();
  }
}
