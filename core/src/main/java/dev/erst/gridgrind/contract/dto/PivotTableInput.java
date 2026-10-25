package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Protocol-facing authored pivot-table definition. */
public record PivotTableInput(
    String name,
    String sheetName,
    Source source,
    Anchor anchor,
    @ProtocolField(optional = true) List<String> rowLabels,
    @ProtocolField(optional = true) List<String> columnLabels,
    @ProtocolField(optional = true) List<String> reportFilters,
    List<DataField> dataFields) {
  /** Reads a pivot-table payload while defaulting omitted axis lists to empty. */
  @JsonCreator
  public static PivotTableInput create(
      @JsonProperty("name") String name,
      @JsonProperty("sheetName") String sheetName,
      @JsonProperty("source") Source source,
      @JsonProperty("anchor") Anchor anchor,
      @JsonProperty("rowLabels") List<String> rowLabels,
      @JsonProperty("columnLabels") List<String> columnLabels,
      @JsonProperty("reportFilters") List<String> reportFilters,
      @JsonProperty("dataFields") List<DataField> dataFields) {
    return new PivotTableInput(
        name,
        sheetName,
        source,
        anchor,
        rowLabels == null ? List.of() : rowLabels,
        columnLabels == null ? List.of() : columnLabels,
        reportFilters == null ? List.of() : reportFilters,
        dataFields);
  }

  public PivotTableInput {
    name = ExcelPivotTableNaming.validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    rowLabels = copyDistinctNames(rowLabels, "rowLabels");
    columnLabels = copyDistinctNames(columnLabels, "columnLabels");
    reportFilters = copyDistinctNames(reportFilters, "reportFilters");
    dataFields = copyDataFields(dataFields);
    requireDisjointAxisAssignments(rowLabels, columnLabels, reportFilters, dataFields);
  }

  /** Supported authored pivot-data source kinds. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Source.Range.class, name = "RANGE"),
    @JsonSubTypes.Type(value = Source.NamedRange.class, name = "NAMED_RANGE"),
    @JsonSubTypes.Type(value = Source.Table.class, name = "TABLE")
  })
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
        name = ProtocolDefinedNameValidation.validateName(name);
      }
    }

    /** Existing workbook-global table source. */
    record Table(String name) implements Source {
      public Table {
        name = ProtocolDefinedNameValidation.validateName(name);
      }
    }
  }

  /** Authored top-left destination anchor for one pivot table. */
  public record Anchor(String topLeftAddress) {
    public Anchor {
      topLeftAddress = ProtocolCellAddressValidation.validateAddress(topLeftAddress);
    }
  }

  /** One authored pivot data field. */
  public record DataField(
      String sourceColumnName,
      ExcelPivotDataConsolidateFunction function,
      @ProtocolField(optional = true) String displayName,
      Optional<String> valueFormat) {
    /** Reads one pivot data-field while defaulting displayName and valueFormat when omitted. */
    @JsonCreator
    public static DataField create(
        @JsonProperty("sourceColumnName") String sourceColumnName,
        @JsonProperty("function") ExcelPivotDataConsolidateFunction function,
        @JsonProperty("displayName") String displayName,
        @JsonProperty("valueFormat") Optional<String> valueFormat) {
      return new DataField(
          sourceColumnName,
          function,
          displayName == null ? sourceColumnName : displayName,
          valueFormat == null ? Optional.empty() : valueFormat);
    }

    public DataField {
      sourceColumnName = requireNonBlank(sourceColumnName, "sourceColumnName");
      Objects.requireNonNull(function, "function must not be null");
      displayName = requireNonBlank(displayName, "displayName");
      Objects.requireNonNull(valueFormat, "valueFormat must not be null");
      if (valueFormat.isPresent() && valueFormat.orElseThrow().isBlank()) {
        throw new IllegalArgumentException("valueFormat must not be blank");
      }
    }
  }

  private static List<String> copyDistinctNames(List<String> names, String fieldName) {
    Objects.requireNonNull(names, fieldName + " must not be null");
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
      addUniqueAxisName(used, name);
    }
    for (String name : columnLabels) {
      addUniqueAxisName(used, name);
    }
    for (String name : reportFilters) {
      addUniqueAxisName(used, name);
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

  private static void addUniqueAxisName(Set<String> used, String name) {
    String key = name.toUpperCase(Locale.ROOT);
    if (!used.add(key)) {
      throw new IllegalArgumentException(
          "pivot source column assignments must be disjoint across rowLabels, columnLabels, and reportFilters: "
              + name);
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
