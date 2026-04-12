package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.util.CellReference;

/** Immutable workbook-core definition of one supported pivot table to create or replace. */
public record ExcelPivotTableDefinition(
    String name,
    String sheetName,
    Source source,
    Anchor anchor,
    List<String> rowLabels,
    List<String> columnLabels,
    List<String> reportFilters,
    List<DataField> dataFields) {

  public ExcelPivotTableDefinition {
    name = validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    rowLabels = copyDistinctNames(rowLabels, "rowLabels");
    columnLabels = copyDistinctNames(columnLabels, "columnLabels");
    reportFilters = copyDistinctNames(reportFilters, "reportFilters");
    dataFields = copyDataFields(dataFields);
    requireDisjointAxisAssignments(rowLabels, columnLabels, reportFilters, dataFields);
  }

  /** Validates and canonicalizes one pivot-table identifier. */
  public static String validateName(String name) {
    return ExcelPivotTableNaming.validateName(name);
  }

  /** Supported authored pivot-data source kinds. */
  public sealed interface Source permits Source.Range, Source.NamedRange, Source.Table {

    /** Explicit contiguous sheet range source with a header row in the first row. */
    record Range(String sheetName, String range) implements Source {
      public Range {
        ExcelSheetNames.requireValid(sheetName, "sheetName");
        Objects.requireNonNull(range, "range must not be null");
        if (range.isBlank()) {
          throw new IllegalArgumentException("range must not be blank");
        }
      }
    }

    /** Existing contiguous workbook or sheet-scoped named range source. */
    record NamedRange(String name) implements Source {
      public NamedRange {
        name = ExcelNamedRangeDefinition.validateName(name);
      }
    }

    /** Existing workbook-global table source. */
    record Table(String name) implements Source {
      public Table {
        name = ExcelTableDefinition.validateName(name);
      }
    }
  }

  /** Authored top-left destination anchor for one pivot table. */
  public record Anchor(String topLeftAddress) {
    public Anchor {
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      if (topLeftAddress.isBlank()) {
        throw new IllegalArgumentException("topLeftAddress must not be blank");
      }
      new CellReference(topLeftAddress);
    }
  }

  /** One authored pivot data field. */
  public record DataField(
      String sourceColumnName,
      ExcelPivotDataConsolidateFunction function,
      String displayName,
      String valueFormat) {
    public DataField {
      sourceColumnName = requireNonBlank(sourceColumnName, "sourceColumnName");
      Objects.requireNonNull(function, "function must not be null");
      displayName = displayName == null || displayName.isBlank() ? sourceColumnName : displayName;
      if (valueFormat != null && valueFormat.isBlank()) {
        throw new IllegalArgumentException("valueFormat must not be blank");
      }
    }
  }

  private static List<String> copyDistinctNames(List<String> names, String fieldName) {
    if (names == null) {
      return List.of();
    }
    List<String> copy = new ArrayList<>(names.size());
    Set<String> unique = new LinkedHashSet<>();
    for (String name : names) {
      String normalized = requireNonBlank(name, fieldName + " value");
      String key = normalized.toUpperCase(Locale.ROOT);
      if (!unique.add(key)) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
      }
      copy.add(normalized);
    }
    return List.copyOf(copy);
  }

  private static List<DataField> copyDataFields(List<DataField> dataFields) {
    Objects.requireNonNull(dataFields, "dataFields must not be null");
    List<DataField> copy = new ArrayList<>(dataFields.size());
    for (DataField dataField : dataFields) {
      copy.add(Objects.requireNonNull(dataField, "dataFields must not contain null values"));
    }
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("dataFields must not be empty");
    }
    return List.copyOf(copy);
  }

  private static void requireDisjointAxisAssignments(
      List<String> rowLabels,
      List<String> columnLabels,
      List<String> reportFilters,
      List<DataField> dataFields) {
    Set<String> used = new LinkedHashSet<>();
    for (String name : rowLabels) {
      addUniqueAxisName(used, name, "rowLabels");
    }
    for (String name : columnLabels) {
      addUniqueAxisName(used, name, "columnLabels");
    }
    for (String name : reportFilters) {
      addUniqueAxisName(used, name, "reportFilters");
    }
    for (DataField dataField : dataFields) {
      String key = dataField.sourceColumnName().toUpperCase(Locale.ROOT);
      if (used.contains(key)) {
        throw new IllegalArgumentException(
            "dataFields must not reuse a source column already assigned to rowLabels, columnLabels, or reportFilters: "
                + dataField.sourceColumnName());
      }
    }
  }

  private static void addUniqueAxisName(Set<String> used, String name, String fieldName) {
    String key = name.toUpperCase(Locale.ROOT);
    if (!used.add(key)) {
      throw new IllegalArgumentException(
          "pivot source column assignments must be disjoint across rowLabels, columnLabels, and reportFilters: "
              + name
              + " in "
              + fieldName);
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
