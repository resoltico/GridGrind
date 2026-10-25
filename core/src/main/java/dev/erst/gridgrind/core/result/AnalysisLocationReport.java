package dev.erst.gridgrind.core.result;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Precise workbook location attached to one derived analysis finding. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = AnalysisLocationReport.Workbook.class, name = "WORKBOOK"),
  @JsonSubTypes.Type(value = AnalysisLocationReport.Sheet.class, name = "SHEET"),
  @JsonSubTypes.Type(value = AnalysisLocationReport.Cell.class, name = "CELL"),
  @JsonSubTypes.Type(value = AnalysisLocationReport.Range.class, name = "RANGE"),
  @JsonSubTypes.Type(value = AnalysisLocationReport.NamedRange.class, name = "NAMED_RANGE")
})
public sealed interface AnalysisLocationReport
    permits AnalysisLocationReport.Workbook,
        AnalysisLocationReport.Sheet,
        AnalysisLocationReport.Cell,
        AnalysisLocationReport.Range,
        AnalysisLocationReport.NamedRange {
  /** Workbook-level finding with no narrower location. */
  record Workbook() implements AnalysisLocationReport {}

  /** One whole-sheet finding. */
  record Sheet(SheetLabel sheetName) implements AnalysisLocationReport {
    public Sheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
    }
  }

  /** One concrete cell finding. */
  record Cell(SheetLabel sheetName, String address) implements AnalysisLocationReport {
    public Cell {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** One rectangular range finding. */
  record Range(SheetLabel sheetName, String range) implements AnalysisLocationReport {
    public Range {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** One named-range finding. */
  record NamedRange(String name, NamedRangeScope scope) implements AnalysisLocationReport {
    public NamedRange {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
    }
  }
}
