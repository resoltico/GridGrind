package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyStrings;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import java.util.List;
import java.util.Objects;

/** Derived surface-description results such as formulas, schema, and named ranges. */
public sealed interface WorkbookSurfaceResult extends WorkbookReadIntrospectionResult
    permits WorkbookSurfaceResult.FormulaSurfaceResult,
        WorkbookSurfaceResult.SheetSchemaResult,
        WorkbookSurfaceResult.NamedRangeSurfaceResult {

  /** Returns grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceResult(String stepId, FormulaSurface analysis)
      implements WorkbookSurfaceResult {
    public FormulaSurfaceResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns inferred schema facts for one rectangular sheet window. */
  record SheetSchemaResult(String stepId, SheetSchema analysis) implements WorkbookSurfaceResult {
    public SheetSchemaResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Returns high-level characterization of selected named ranges. */
  record NamedRangeSurfaceResult(String stepId, NamedRangeSurface analysis)
      implements WorkbookSurfaceResult {
    public NamedRangeSurfaceResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(analysis, "analysis must not be null");
    }
  }

  /** Grouped formula usage facts across one or more sheets. */
  record FormulaSurface(int totalFormulaCellCount, List<SheetFormulaSurface> sheets) {
    public FormulaSurface {
      if (totalFormulaCellCount < 0) {
        throw new IllegalArgumentException("totalFormulaCellCount must not be negative");
      }
      sheets = copyValues(sheets, "sheets");
    }
  }

  /** Formula usage facts for one sheet. */
  record SheetFormulaSurface(
      String sheetName,
      int formulaCellCount,
      int distinctFormulaCount,
      List<FormulaPattern> formulas) {
    public SheetFormulaSurface {
      sheetName = requireNonBlank(sheetName, "sheetName");
      if (formulaCellCount < 0) {
        throw new IllegalArgumentException("formulaCellCount must not be negative");
      }
      if (distinctFormulaCount < 0) {
        throw new IllegalArgumentException("distinctFormulaCount must not be negative");
      }
      formulas = copyValues(formulas, "formulas");
    }
  }

  /** One grouped formula pattern and the addresses where it appears. */
  record FormulaPattern(String formula, int occurrenceCount, List<String> addresses) {
    public FormulaPattern {
      formula = requireNonBlank(formula, "formula");
      if (occurrenceCount <= 0) {
        throw new IllegalArgumentException("occurrenceCount must be greater than 0");
      }
      addresses = copyStrings(addresses, "addresses");
    }
  }

  /** Inferred schema facts for one rectangular sheet window. */
  record SheetSchema(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      int dataRowCount,
      List<SchemaColumn> columns) {
    public SheetSchema {
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      if (rowCount <= 0) {
        throw new IllegalArgumentException("rowCount must be greater than 0");
      }
      if (columnCount <= 0) {
        throw new IllegalArgumentException("columnCount must be greater than 0");
      }
      if (dataRowCount < 0) {
        throw new IllegalArgumentException("dataRowCount must not be negative");
      }
      columns = copyValues(columns, "columns");
    }
  }

  /** One inferred schema column with header text and observed value-type counts. */
  record SchemaColumn(
      int columnIndex,
      String columnAddress,
      String headerDisplayValue,
      int populatedCellCount,
      int blankCellCount,
      List<TypeCount> observedTypes,
      String dominantType) {
    public SchemaColumn {
      if (columnIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("columnIndex", columnIndex));
      }
      columnAddress = requireNonBlank(columnAddress, "columnAddress");
      Objects.requireNonNull(headerDisplayValue, "headerDisplayValue must not be null");
      if (populatedCellCount < 0) {
        throw new IllegalArgumentException("populatedCellCount must not be negative");
      }
      if (blankCellCount < 0) {
        throw new IllegalArgumentException("blankCellCount must not be negative");
      }
      observedTypes = copyValues(observedTypes, "observedTypes");
    }
  }

  /** Count of one observed cell type inside a schema column. */
  record TypeCount(String type, int count) {
    public TypeCount {
      type = requireNonBlank(type, "type");
      if (count <= 0) {
        throw new IllegalArgumentException("count must be greater than 0");
      }
    }
  }

  /** High-level characterization of selected named ranges. */
  record NamedRangeSurface(
      int workbookScopedCount,
      int sheetScopedCount,
      int rangeBackedCount,
      int formulaBackedCount,
      List<NamedRangeSurfaceEntry> namedRanges) {
    public NamedRangeSurface {
      if (workbookScopedCount < 0) {
        throw new IllegalArgumentException("workbookScopedCount must not be negative");
      }
      if (sheetScopedCount < 0) {
        throw new IllegalArgumentException("sheetScopedCount must not be negative");
      }
      if (rangeBackedCount < 0) {
        throw new IllegalArgumentException("rangeBackedCount must not be negative");
      }
      if (formulaBackedCount < 0) {
        throw new IllegalArgumentException("formulaBackedCount must not be negative");
      }
      namedRanges = copyValues(namedRanges, "namedRanges");
    }
  }

  /** One named-range surface entry classified by scope and backing kind. */
  record NamedRangeSurfaceEntry(
      String name, ExcelNamedRangeScope scope, String refersToFormula, NamedRangeBackingKind kind) {
    public NamedRangeSurfaceEntry {
      name = requireNonBlank(name, "name");
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      Objects.requireNonNull(kind, "kind must not be null");
    }
  }

  /** Distinguishes range-backed named ranges from formula-backed named ranges. */
  enum NamedRangeBackingKind {
    RANGE,
    FORMULA
  }
}
