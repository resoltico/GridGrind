package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual report for one table loaded from a workbook. */
public record TableEntryReport(
    String name,
    String sheetName,
    String range,
    int headerRowCount,
    int totalsRowCount,
    List<String> columnNames,
    List<TableColumnReport> columns,
    TableStyleReport style,
    boolean hasAutofilter,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> comment,
    boolean published,
    boolean insertRow,
    boolean insertRowShift,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> headerRowCellStyle,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> dataCellStyle,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> totalsRowCellStyle) {
  /** Creates a table report with defaulted per-column metadata and optional flags. */
  public TableEntryReport(
      String name,
      String sheetName,
      String range,
      int headerRowCount,
      int totalsRowCount,
      List<String> columnNames,
      TableStyleReport style,
      boolean hasAutofilter) {
    this(
        name,
        sheetName,
        range,
        headerRowCount,
        totalsRowCount,
        columnNames,
        columnNames.stream().map(columnName -> new TableColumnReport(0L, columnName)).toList(),
        style,
        hasAutofilter,
        Optional.empty(),
        false,
        false,
        false,
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  public TableEntryReport {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    Objects.requireNonNull(range, "range must not be null");
    if (headerRowCount < 0) {
      throw new IllegalArgumentException("headerRowCount must not be negative");
    }
    if (totalsRowCount < 0) {
      throw new IllegalArgumentException("totalsRowCount must not be negative");
    }
    Objects.requireNonNull(columnNames, "columnNames must not be null");
    columnNames = copyColumnNames(columnNames);
    Objects.requireNonNull(columns, "columns must not be null");
    columns = copyColumns(columns);
    comment = normalizeOptional(comment);
    headerRowCellStyle = normalizeOptional(headerRowCellStyle);
    dataCellStyle = normalizeOptional(dataCellStyle);
    totalsRowCellStyle = normalizeOptional(totalsRowCellStyle);
    Objects.requireNonNull(style, "style must not be null");
  }

  @JsonCreator
  @SuppressWarnings("PMD.ExcessiveParameterList")
  static TableEntryReport create(
      @JsonProperty("name") String name,
      @JsonProperty("sheetName") String sheetName,
      @JsonProperty("range") String range,
      @JsonProperty("headerRowCount") int headerRowCount,
      @JsonProperty("totalsRowCount") int totalsRowCount,
      @JsonProperty("columnNames") List<String> columnNames,
      @JsonProperty("columns") List<TableColumnReport> columns,
      @JsonProperty("style") TableStyleReport style,
      @JsonProperty("hasAutofilter") boolean hasAutofilter,
      @JsonProperty("comment") Optional<String> comment,
      @JsonProperty("published") boolean published,
      @JsonProperty("insertRow") boolean insertRow,
      @JsonProperty("insertRowShift") boolean insertRowShift,
      @JsonProperty("headerRowCellStyle") Optional<String> headerRowCellStyle,
      @JsonProperty("dataCellStyle") Optional<String> dataCellStyle,
      @JsonProperty("totalsRowCellStyle") Optional<String> totalsRowCellStyle) {
    return new TableEntryReport(
        name,
        sheetName,
        range,
        headerRowCount,
        totalsRowCount,
        columnNames,
        columns,
        style,
        hasAutofilter,
        comment,
        published,
        insertRow,
        insertRowShift,
        headerRowCellStyle,
        dataCellStyle,
        totalsRowCellStyle);
  }

  private static List<String> copyColumnNames(List<String> columnNames) {
    for (String columnName : columnNames) {
      Objects.requireNonNull(columnName, "columnNames must not contain nulls");
    }
    return List.copyOf(columnNames);
  }

  private static List<TableColumnReport> copyColumns(List<TableColumnReport> columns) {
    for (TableColumnReport column : columns) {
      Objects.requireNonNull(column, "columns must not contain nulls");
    }
    return List.copyOf(columns);
  }

  private static Optional<String> normalizeOptional(Optional<String> value) {
    Optional<String> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    if (normalized.isPresent()) {
      String text = normalized.orElseThrow();
      if (text.isBlank()) {
        return Optional.empty();
      }
      return Optional.of(text);
    }
    return Optional.empty();
  }
}
