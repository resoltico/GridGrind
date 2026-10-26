package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.TableLabel;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual report for one table loaded from a workbook. */
@CatalogType(
    id = "TableEntryReport",
    summary = "Exact workbook table report used by table-facts assertions.",
    displayOrder = 91)
public record TableEntryReport(
    TableLabel name,
    SheetLabel sheetName,
    RangeLabel range,
    Structure structure,
    TableStyleReport style,
    @WireDefault(factory = Behavior.class, method = "defaults") Behavior behavior,
    @WireDefault(factory = Presentation.class, method = "empty") Presentation presentation) {
  /** Creates a table report with defaulted per-column metadata and optional presentation state. */
  public TableEntryReport(
      TableLabel name,
      SheetLabel sheetName,
      RangeLabel range,
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
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(structure, "structure must not be null");
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(behavior, "behavior must not be null");
    Objects.requireNonNull(presentation, "presentation must not be null");
  }

  /** Structural table facts that must move together to describe one persisted table shape. */
  @CatalogType(
      id = "TableEntryStructureReport",
      summary = "Structural table facts including header/totals counts and per-column metadata.",
      displayOrder = 92)
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
  @CatalogType(
      id = "TableEntryBehaviorReport",
      summary = "Persisted workbook-table behavior toggles.",
      displayOrder = 93)
  public record Behavior(
      boolean hasAutofilter, boolean published, boolean insertRow, boolean insertRowShift) {
    /** Returns the standard persisted table behavior. */
    public static Behavior defaults() {
      return new Behavior(false, false, false, false);
    }
  }

  /** Optional authored notes and style labels attached to one persisted workbook table. */
  @CatalogType(
      id = "TableEntryPresentationReport",
      summary = "Optional table comment and style labels attached to one persisted workbook table.",
      displayOrder = 94)
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

    /** Returns absent optional table presentation facts. */
    public static Presentation empty() {
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
