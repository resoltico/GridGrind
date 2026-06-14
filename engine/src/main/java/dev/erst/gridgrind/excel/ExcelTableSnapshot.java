package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual table metadata loaded from a workbook. */
public record ExcelTableSnapshot(
    String name,
    String sheetName,
    String range,
    Structure structure,
    ExcelTableStyleSnapshot style,
    Behavior behavior,
    Presentation presentation) {
  /** Creates a table snapshot with derived per-column metadata and empty presentation state. */
  public ExcelTableSnapshot(
      String name,
      String sheetName,
      String range,
      int headerRowCount,
      int totalsRowCount,
      List<String> columnNames,
      ExcelTableStyleSnapshot style,
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

  public ExcelTableSnapshot {
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

  /** Structural table facts that move together when the workbook table shape changes. */
  public record Structure(
      int headerRowCount,
      int totalsRowCount,
      List<String> columnNames,
      List<ExcelTableColumnSnapshot> columns) {
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
              .map(columnName -> new ExcelTableColumnSnapshot(0L, columnName, "", "", "", ""))
              .toList());
    }
  }

  /** Persisted workbook-table behavior toggles. */
  public record Behavior(
      boolean hasAutofilter, boolean published, boolean insertRow, boolean insertRowShift) {}

  /** Optional workbook-table comment and style labels. */
  public record Presentation(
      Optional<String> comment,
      Optional<String> headerRowCellStyle,
      Optional<String> dataCellStyle,
      Optional<String> totalsRowCellStyle) {
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
    List<String> copied = List.copyOf(columnNames);
    for (String columnName : copied) {
      Objects.requireNonNull(columnName, "columnNames must not contain nulls");
    }
    return copied;
  }

  private static List<ExcelTableColumnSnapshot> copyColumns(
      List<ExcelTableColumnSnapshot> columns) {
    Objects.requireNonNull(columns, "columns must not be null");
    List<ExcelTableColumnSnapshot> copied = List.copyOf(columns);
    for (ExcelTableColumnSnapshot column : copied) {
      Objects.requireNonNull(column, "columns must not contain nulls");
    }
    return copied;
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
