package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.FormulaSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceReport;
import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Objects;

/** Derived surface-description results such as formulas, schema, and named ranges. */
public sealed interface WorkbookSurfaceResult extends WorkbookReadSurfaceResult
    permits WorkbookSurfaceResult.FormulaSurfaceResult,
        WorkbookSurfaceResult.SheetSchemaResult,
        WorkbookSurfaceResult.NamedRangeSurfaceResult {

  /** Returns grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceResult(StepId stepId, FormulaSurfaceReport surface)
      implements WorkbookSurfaceResult {
    public FormulaSurfaceResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }

  /** Returns inferred schema facts for one rectangular sheet window. */
  record SheetSchemaResult(StepId stepId, SheetSchemaReport surface)
      implements WorkbookSurfaceResult {
    public SheetSchemaResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }

  /** Returns high-level characterization of selected named ranges. */
  record NamedRangeSurfaceResult(StepId stepId, NamedRangeSurfaceReport surface)
      implements WorkbookSurfaceResult {
    public NamedRangeSurfaceResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(surface, "surface must not be null");
    }
  }
}
