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
    Structure structure,
    TableStyleReport style,
    Behavior behavior,
    Presentation presentation) {
  /** Creates a table report with defaulted per-column metadata and optional presentation state. */
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
        Structure.withDerivedColumns(headerRowCount, totalsRowCount, columnNames),
        style,
        new Behavior(hasAutofilter, false, false, false),
        Presentation.empty());
  }

  public TableEntryReport {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(structure, "structure must not be null");
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(behavior, "behavior must not be null");
    Objects.requireNonNull(presentation, "presentation must not be null");
  }

  @JsonCreator
  static TableEntryReport create(
      @JsonProperty("name") String name,
      @JsonProperty("sheetName") String sheetName,
      @JsonProperty("range") String range,
      @JsonProperty("structure") Structure structure,
      @JsonProperty("style") TableStyleReport style,
      @JsonProperty("behavior") Behavior behavior,
      @JsonProperty("presentation") Presentation presentation) {
    return new TableEntryReport(
        name,
        sheetName,
        range,
        Objects.requireNonNull(structure, "structure must not be null"),
        style,
        Objects.requireNonNullElseGet(behavior, Behavior::defaults),
        Objects.requireNonNullElseGet(presentation, Presentation::empty));
  }

  /** Structural table facts that must move together to describe one persisted table shape. */
  public record Structure(
      int headerRowCount,
      int totalsRowCount,
      List<String> columnNames,
      List<TableColumnReport> columns) {
    public Structure {
      if (headerRowCount < 0) {
        throw new IllegalArgumentException("headerRowCount must not be negative");
      }
      if (totalsRowCount < 0) {
        throw new IllegalArgumentException("totalsRowCount must not be negative");
      }
      columnNames = copyColumnNames(columnNames);
      columns = copyColumns(columns);
    }

    static Structure withDerivedColumns(
        int headerRowCount, int totalsRowCount, List<String> columnNames) {
      List<String> copiedColumnNames = copyColumnNames(columnNames);
      return new Structure(
          headerRowCount,
          totalsRowCount,
          copiedColumnNames,
          copiedColumnNames.stream()
              .map(columnName -> new TableColumnReport(0L, columnName))
              .toList());
    }
  }

  /** Persisted workbook-table behavior toggles. */
  public record Behavior(
      boolean hasAutofilter, boolean published, boolean insertRow, boolean insertRowShift) {
    static Behavior defaults() {
      return new Behavior(false, false, false, false);
    }
  }

  /** Optional authored notes and style labels attached to one persisted workbook table. */
  public record Presentation(
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> comment,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> headerRowCellStyle,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> dataCellStyle,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> totalsRowCellStyle) {
    public Presentation {
      comment = normalizeOptional(comment);
      headerRowCellStyle = normalizeOptional(headerRowCellStyle);
      dataCellStyle = normalizeOptional(dataCellStyle);
      totalsRowCellStyle = normalizeOptional(totalsRowCellStyle);
    }

    static Presentation empty() {
      return new Presentation(
          Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty());
    }
  }

  private static List<String> copyColumnNames(List<String> columnNames) {
    Objects.requireNonNull(columnNames, "columnNames must not be null");
    for (String columnName : columnNames) {
      Objects.requireNonNull(columnName, "columnNames must not contain nulls");
    }
    return List.copyOf(columnNames);
  }

  private static List<TableColumnReport> copyColumns(List<TableColumnReport> columns) {
    Objects.requireNonNull(columns, "columns must not be null");
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
