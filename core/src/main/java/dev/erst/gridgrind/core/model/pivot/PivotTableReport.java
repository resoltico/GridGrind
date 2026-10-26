package dev.erst.gridgrind.core.model.pivot;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotTableNaming;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual pivot-table report returned by pivot-table reads. */
@Union(discriminator = "type", catalog = true, displayOrder = 9)
public sealed interface PivotTableReport
    permits dev.erst.gridgrind.core.model.pivot.PivotTableReport.Supported,
        dev.erst.gridgrind.core.model.pivot.PivotTableReport.Unsupported {

  /** Workbook-global pivot-table name. */
  String name();

  /** Sheet that owns the pivot table relation. */
  SheetName sheetName();

  /** Stored top-left anchor and current persisted location range. */
  Anchor anchor();

  /** Supported pivot-table report with first-class factual structure. */
  @Variant(id = "SUPPORTED", summary = "Exact supported pivot-table report.")
  record Supported(
      String name,
      SheetName sheetName,
      Anchor anchor,
      PivotSourceReport source,
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
  @Variant(
      id = "UNSUPPORTED",
      summary = "Exact unsupported pivot-table report preserved from the workbook.")
  record Unsupported(String name, SheetName sheetName, Anchor anchor, String detail)
      implements PivotTableReport {
    public Unsupported {
      validateCommon(name, sheetName, anchor);
      detail = requireNonBlank(detail, "detail");
    }
  }

  /** Factual persisted pivot source. */
  @Union(discriminator = "type", catalog = true, displayOrder = 10)
  sealed interface PivotSourceReport
      permits dev.erst.gridgrind.core.model.pivot.PivotTableReport.PivotSourceReport.Range,
          dev.erst.gridgrind.core.model.pivot.PivotTableReport.PivotSourceReport.NamedRange,
          dev.erst.gridgrind.core.model.pivot.PivotTableReport.PivotSourceReport.Table {

    /** Sheet range source stored directly in the pivot cache definition. */
    @Variant(id = "RANGE", summary = "Pivot source resolved from one sheet range.")
    record Range(SheetName sheetName, RangeLabel range) implements PivotSourceReport {
      public Range {
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
        java.util.Objects.requireNonNull(range, "range must not be null");
      }
    }

    /** Named-range-backed source plus the currently resolved source range. */
    @Variant(id = "NAMED_RANGE", summary = "Pivot source resolved from one named range.")
    record NamedRange(DefinedName name, SheetName sheetName, RangeLabel range)
        implements PivotSourceReport {
      public NamedRange {
        java.util.Objects.requireNonNull(name, "name must not be null");
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
        java.util.Objects.requireNonNull(range, "range must not be null");
      }
    }

    /** Table-backed source plus the currently resolved table range. */
    @Variant(id = "TABLE", summary = "Pivot source resolved from one workbook table.")
    record Table(TableName name, SheetName sheetName, RangeLabel range)
        implements PivotSourceReport {
      public Table {
        java.util.Objects.requireNonNull(name, "name must not be null");
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
        java.util.Objects.requireNonNull(range, "range must not be null");
      }
    }
  }

  /** Factual stored destination anchor and persisted location range. */
  @CatalogType(
      id = "PivotTableAnchorReport",
      summary = "Exact pivot-table anchor report.",
      displayOrder = 97)
  record Anchor(CellAddress topLeftAddress, RangeLabel locationRange) {
    public Anchor {
      java.util.Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      java.util.Objects.requireNonNull(locationRange, "locationRange must not be null");
    }
  }

  /** Factual pivot field bound to one source column. */
  @CatalogType(
      id = "PivotTableFieldReport",
      summary = "Exact pivot field report bound to one source column.",
      displayOrder = 98)
  record Field(int sourceColumnIndex, String sourceColumnName) {
    public Field {
      if (sourceColumnIndex < 0) {
        throw new IllegalArgumentException("sourceColumnIndex must not be negative");
      }
      sourceColumnName = requireNonBlank(sourceColumnName, "sourceColumnName");
    }
  }

  /** Factual pivot data field. */
  @CatalogType(
      id = "PivotTableDataFieldReport",
      summary = "Exact pivot data-field report.",
      displayOrder = 99)
  record DataField(
      int sourceColumnIndex,
      String sourceColumnName,
      ExcelPivotDataConsolidateFunction function,
      String displayName,
      Optional<String> valueFormat) {
    public DataField {
      if (sourceColumnIndex < 0) {
        throw new IllegalArgumentException("sourceColumnIndex must not be negative");
      }
      sourceColumnName = requireNonBlank(sourceColumnName, "sourceColumnName");
      Objects.requireNonNull(function, "function must not be null");
      displayName = requireNonBlank(displayName, "displayName");
      Objects.requireNonNull(valueFormat, "valueFormat must not be null");
      if (valueFormat.isPresent() && valueFormat.orElseThrow().isBlank()) {
        throw new IllegalArgumentException("valueFormat must not be blank");
      }
    }
  }

  private static void validateCommon(String name, SheetName sheetName, Anchor anchor) {
    ExcelPivotTableNaming.validateName(name);
    java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
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
