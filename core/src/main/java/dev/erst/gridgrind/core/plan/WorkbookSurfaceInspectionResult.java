package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceReport;
import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Objects;

/** Derived workbook-surface summaries that remain factual reads rather than health analysis. */
public sealed interface WorkbookSurfaceInspectionResult extends InspectionSurfaceResult
    permits WorkbookSurfaceInspectionResult.FormulaSurfaceResult,
        WorkbookSurfaceInspectionResult.SheetSchemaResult,
        WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult {

  /** Returns grouped formula usage facts across one or more sheets. */
  @Variant(
      id = "GET_FORMULA_SURFACE",
      summary = "Returns grouped formula usage facts across one or more sheets.")
  record FormulaSurfaceResult(StepId stepId, FormulaSurfaceReport surface)
      implements WorkbookSurfaceInspectionResult {
    public FormulaSurfaceResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }

  /** Returns inferred schema facts for one sheet window. */
  @Variant(id = "GET_SHEET_SCHEMA", summary = "Returns inferred schema facts for one sheet window.")
  record SheetSchemaResult(StepId stepId, SheetSchemaReport surface)
      implements WorkbookSurfaceInspectionResult {
    public SheetSchemaResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }

  /** Returns high-level characterization of named ranges. */
  @Variant(
      id = "GET_NAMED_RANGE_SURFACE",
      summary = "Returns high-level characterization of named ranges.")
  record NamedRangeSurfaceResult(StepId stepId, NamedRangeSurfaceReport surface)
      implements WorkbookSurfaceInspectionResult {
    public NamedRangeSurfaceResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }
}
