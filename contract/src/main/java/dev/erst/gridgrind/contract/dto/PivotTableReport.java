package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.ExcelSheetNames;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Protocol-facing factual pivot-table report returned by pivot-table reads. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = PivotTableReport.Supported.class, name = "SUPPORTED"),
  @JsonSubTypes.Type(value = PivotTableReport.Unsupported.class, name = "UNSUPPORTED")
})
public sealed interface PivotTableReport
    permits PivotTableReport.Supported, PivotTableReport.Unsupported {

  /** Workbook-global pivot-table name. */
  String name();

  /** Sheet that owns the pivot table relation. */
  String sheetName();

  /** Stored top-left anchor and current persisted location range. */
  Anchor anchor();

  /** Supported pivot-table report with first-class factual structure. */
  record Supported(
      String name,
      String sheetName,
      Anchor anchor,
      Source source,
      List<Field> rowLabels,
      List<Field> columnLabels,
      List<Field> reportFilters,
      List<DataField> dataFields,
      boolean valuesAxisOnColumns)
      implements PivotTableReport {
    public Supported {
      validateCommon(name, sheetName, anchor);
      Objects.requireNonNull(source, "source must not be null");
      rowLabels = copyValues(rowLabels, "rowLabels");
      columnLabels = copyValues(columnLabels, "columnLabels");
      reportFilters = copyValues(reportFilters, "reportFilters");
      dataFields = copyValues(dataFields, "dataFields");
    }
  }

  /** Unsupported or malformed pivot-table report preserved for truthful readback. */
  record Unsupported(String name, String sheetName, Anchor anchor, String detail)
      implements PivotTableReport {
    public Unsupported {
      validateCommon(name, sheetName, anchor);
      detail = requireNonBlank(detail, "detail");
    }
  }

  /** Factual persisted pivot source. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Source.Range.class, name = "RANGE"),
    @JsonSubTypes.Type(value = Source.NamedRange.class, name = "NAMED_RANGE"),
    @JsonSubTypes.Type(value = Source.Table.class, name = "TABLE")
  })
  sealed interface Source permits Source.Range, Source.NamedRange, Source.Table {

    /** Sheet range source stored directly in the pivot cache definition. */
    record Range(String sheetName, String range) implements Source {
      public Range {
        ExcelSheetNames.requireValid(sheetName, "sheetName");
        range = requireNonBlank(range, "range");
      }
    }

    /** Named-range-backed source plus the currently resolved source range. */
    record NamedRange(String name, String sheetName, String range) implements Source {
      public NamedRange {
        name = ProtocolDefinedNameValidation.validateName(name);
        ExcelSheetNames.requireValid(sheetName, "sheetName");
        range = requireNonBlank(range, "range");
      }
    }

    /** Table-backed source plus the currently resolved table range. */
    record Table(String name, String sheetName, String range) implements Source {
      public Table {
        name = ProtocolDefinedNameValidation.validateName(name);
        ExcelSheetNames.requireValid(sheetName, "sheetName");
        range = requireNonBlank(range, "range");
      }
    }
  }

  /** Factual stored destination anchor and persisted location range. */
  record Anchor(String topLeftAddress, String locationRange) {
    public Anchor {
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      locationRange = requireNonBlank(locationRange, "locationRange");
    }
  }

  /** Factual pivot field bound to one source column. */
  record Field(int sourceColumnIndex, String sourceColumnName) {
    public Field {
      if (sourceColumnIndex < 0) {
        throw new IllegalArgumentException("sourceColumnIndex must not be negative");
      }
      sourceColumnName = requireNonBlank(sourceColumnName, "sourceColumnName");
    }
  }

  /** Factual pivot data field. */
  record DataField(
      int sourceColumnIndex,
      String sourceColumnName,
      ExcelPivotDataConsolidateFunction function,
      String displayName,
      String valueFormat) {
    public DataField {
      if (sourceColumnIndex < 0) {
        throw new IllegalArgumentException("sourceColumnIndex must not be negative");
      }
      sourceColumnName = requireNonBlank(sourceColumnName, "sourceColumnName");
      Objects.requireNonNull(function, "function must not be null");
      displayName = requireNonBlank(displayName, "displayName");
      if (valueFormat != null && valueFormat.isBlank()) {
        throw new IllegalArgumentException("valueFormat must not be blank");
      }
    }
  }

  private static void validateCommon(String name, String sheetName, Anchor anchor) {
    ExcelPivotTableNaming.validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(anchor, "anchor must not be null");
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copy);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
