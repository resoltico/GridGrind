package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;
import java.util.Optional;

/** Factual metadata for one persisted workbook table column. */
public record TableColumnReport(
    long id,
    String name,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> uniqueName,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> totalsRowLabel,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> totalsRowFunction,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> calculatedColumnFormula) {
  /** Creates a table-column report with no auxiliary metadata. */
  public TableColumnReport(long id, String name) {
    this(id, name, Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty());
  }

  public TableColumnReport {
    if (id < 0L) {
      throw new IllegalArgumentException("id must not be negative");
    }
    Objects.requireNonNull(name, "name must not be null");
    uniqueName = normalizeOptional(uniqueName);
    totalsRowLabel = normalizeOptional(totalsRowLabel);
    totalsRowFunction = normalizeOptional(totalsRowFunction);
    calculatedColumnFormula = normalizeOptional(calculatedColumnFormula);
  }

  @JsonCreator
  static TableColumnReport create(
      @JsonProperty("id") long id,
      @JsonProperty("name") String name,
      @JsonProperty("uniqueName") Optional<String> uniqueName,
      @JsonProperty("totalsRowLabel") Optional<String> totalsRowLabel,
      @JsonProperty("totalsRowFunction") Optional<String> totalsRowFunction,
      @JsonProperty("calculatedColumnFormula") Optional<String> calculatedColumnFormula) {
    return new TableColumnReport(
        id, name, uniqueName, totalsRowLabel, totalsRowFunction, calculatedColumnFormula);
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
