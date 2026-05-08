package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.dto.FormulaSurfaceReport;
import dev.erst.gridgrind.contract.dto.NamedRangeSurfaceReport;
import dev.erst.gridgrind.contract.dto.SheetSchemaReport;
import java.util.Objects;

/** Derived workbook-surface summaries that remain factual reads rather than health analysis. */
public sealed interface WorkbookSurfaceInspectionResult extends InspectionSurfaceResult
    permits WorkbookSurfaceInspectionResult.FormulaSurfaceResult,
        WorkbookSurfaceInspectionResult.SheetSchemaResult,
        WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult {

  /** Returns grouped formula usage facts across one or more sheets. */
  @JsonTypeName("GET_FORMULA_SURFACE")
  record FormulaSurfaceResult(String stepId, FormulaSurfaceReport surface)
      implements WorkbookSurfaceInspectionResult {
    public FormulaSurfaceResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }

  /** Returns inferred schema facts for one sheet window. */
  @JsonTypeName("GET_SHEET_SCHEMA")
  record SheetSchemaResult(String stepId, SheetSchemaReport surface)
      implements WorkbookSurfaceInspectionResult {
    public SheetSchemaResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }

  /** Returns high-level characterization of named ranges. */
  @JsonTypeName("GET_NAMED_RANGE_SURFACE")
  record NamedRangeSurfaceResult(String stepId, NamedRangeSurfaceReport surface)
      implements WorkbookSurfaceInspectionResult {
    public NamedRangeSurfaceResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }
}
