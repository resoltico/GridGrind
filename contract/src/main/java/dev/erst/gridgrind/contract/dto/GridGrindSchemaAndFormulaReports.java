package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Formula, schema, and named-range surface report families. */
public interface GridGrindSchemaAndFormulaReports {
  /** Grouped formula usage facts across one or more sheets. */
  record FormulaSurfaceReport(int totalFormulaCellCount, List<SheetFormulaSurfaceReport> sheets) {
    public FormulaSurfaceReport {
      sheets = GridGrindResponseSupport.copyValues(sheets, "sheets");
      if (totalFormulaCellCount < 0) {
        throw new IllegalArgumentException("totalFormulaCellCount must not be negative");
      }
    }
  }

  /** Formula usage facts for one sheet. */
  record SheetFormulaSurfaceReport(
      String sheetName,
      int formulaCellCount,
      int distinctFormulaCount,
      List<FormulaPatternReport> formulas) {
    public SheetFormulaSurfaceReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (formulaCellCount < 0) {
        throw new IllegalArgumentException("formulaCellCount must not be negative");
      }
      if (distinctFormulaCount < 0) {
        throw new IllegalArgumentException("distinctFormulaCount must not be negative");
      }
      formulas = GridGrindResponseSupport.copyValues(formulas, "formulas");
    }
  }

  /** One grouped formula pattern and the addresses where it appears. */
  record FormulaPatternReport(String formula, int occurrenceCount, List<String> addresses) {
    public FormulaPatternReport {
      Objects.requireNonNull(formula, "formula must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
      if (occurrenceCount <= 0) {
        throw new IllegalArgumentException("occurrenceCount must be greater than 0");
      }
      addresses = GridGrindResponseSupport.copyStrings(addresses, "addresses");
    }
  }

  /** Inferred schema facts for one rectangular sheet window. */
  record SheetSchemaReport(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      int dataRowCount,
      List<SchemaColumnReport> columns) {
    public SheetSchemaReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (topLeftAddress.isBlank()) {
        throw new IllegalArgumentException("topLeftAddress must not be blank");
      }
      if (rowCount <= 0) {
        throw new IllegalArgumentException("rowCount must be greater than 0");
      }
      if (columnCount <= 0) {
        throw new IllegalArgumentException("columnCount must be greater than 0");
      }
      if (dataRowCount < 0) {
        throw new IllegalArgumentException("dataRowCount must not be negative");
      }
      columns = GridGrindResponseSupport.copyValues(columns, "columns");
    }
  }

  /** One inferred schema column with header text and observed value-type counts. */
  record SchemaColumnReport(
      int columnIndex,
      String columnAddress,
      String headerDisplayValue,
      int populatedCellCount,
      int blankCellCount,
      List<TypeCountReport> observedTypes,
      String dominantType) {
    public SchemaColumnReport {
      if (columnIndex < 0) {
        throw new IllegalArgumentException("columnIndex must not be negative");
      }
      Objects.requireNonNull(columnAddress, "columnAddress must not be null");
      Objects.requireNonNull(headerDisplayValue, "headerDisplayValue must not be null");
      if (columnAddress.isBlank()) {
        throw new IllegalArgumentException("columnAddress must not be blank");
      }
      if (populatedCellCount < 0) {
        throw new IllegalArgumentException("populatedCellCount must not be negative");
      }
      if (blankCellCount < 0) {
        throw new IllegalArgumentException("blankCellCount must not be negative");
      }
      observedTypes = GridGrindResponseSupport.copyValues(observedTypes, "observedTypes");
    }
  }

  /** Count of one observed cell value type inside a schema column. */
  record TypeCountReport(String type, int count) {
    public TypeCountReport {
      Objects.requireNonNull(type, "type must not be null");
      if (type.isBlank()) {
        throw new IllegalArgumentException("type must not be blank");
      }
      if (count <= 0) {
        throw new IllegalArgumentException("count must be greater than 0");
      }
    }
  }

  /** High-level characterization of the named ranges selected by one analysis read. */
  record NamedRangeSurfaceReport(
      int workbookScopedCount,
      int sheetScopedCount,
      int rangeBackedCount,
      int formulaBackedCount,
      List<NamedRangeSurfaceEntryReport> namedRanges) {
    public NamedRangeSurfaceReport {
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
      namedRanges = GridGrindResponseSupport.copyValues(namedRanges, "namedRanges");
    }
  }

  /** One named-range surface entry classified by scope and backing kind. */
  record NamedRangeSurfaceEntryReport(
      String name, NamedRangeScope scope, String refersToFormula, NamedRangeBackingKind kind) {
    public NamedRangeSurfaceEntryReport {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(scope, "scope must not be null");
      Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
      Objects.requireNonNull(kind, "kind must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
    }
  }

  /** Distinguishes range-backed named ranges from formula-backed named ranges. */
  enum NamedRangeBackingKind {
    RANGE,
    FORMULA
  }
}
