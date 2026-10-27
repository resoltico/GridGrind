package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Precise workbook location attached to one derived analysis finding. */
@Union(discriminator = "type")
public sealed interface AnalysisLocationReport
    permits AnalysisLocationReport.Workbook,
        AnalysisLocationReport.Sheet,
        AnalysisLocationReport.Cell,
        AnalysisLocationReport.Range,
        AnalysisLocationReport.NamedRange {
  /** Workbook-level finding with no narrower location. */
  @Variant(id = "WORKBOOK", summary = "Workbook-level finding with no narrower location.")
  record Workbook() implements AnalysisLocationReport {}

  /** One whole-sheet finding. */
  @Variant(id = "SHEET", summary = "One whole-sheet finding.")
  record Sheet(SheetLabel sheetName) implements AnalysisLocationReport {
    public Sheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
    }
  }

  /** One concrete cell finding. */
  @Variant(id = "CELL", summary = "One concrete cell finding.")
  record Cell(SheetLabel sheetName, String address) implements AnalysisLocationReport {
    public Cell {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      address = TextValidation.requireNonBlank(address, "address");
    }
  }

  /** One rectangular range finding. */
  @Variant(id = "RANGE", summary = "One rectangular range finding.")
  record Range(SheetLabel sheetName, String range) implements AnalysisLocationReport {
    public Range {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      range = TextValidation.requireNonBlank(range, "range");
    }
  }

  /** One named-range finding. */
  @Variant(id = "NAMED_RANGE", summary = "One named-range finding.")
  record NamedRange(String name, NamedRangeScopeReport scope) implements AnalysisLocationReport {
    public NamedRange {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      TextValidation.requireNonBlank(name, "name");
    }
  }
}
