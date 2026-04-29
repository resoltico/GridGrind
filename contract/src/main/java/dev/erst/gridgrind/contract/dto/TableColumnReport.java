package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/** Factual metadata for one persisted workbook table column. */
public record TableColumnReport(
    long id,
    String name,
    String uniqueName,
    String totalsRowLabel,
    String totalsRowFunction,
    String calculatedColumnFormula) {
  /** Creates a table-column report with no auxiliary metadata. */
  public TableColumnReport(long id, String name) {
    this(id, name, "", "", "", "");
  }

  public TableColumnReport {
    if (id < 0L) {
      throw new IllegalArgumentException("id must not be negative");
    }
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(uniqueName, "uniqueName must not be null");
    Objects.requireNonNull(totalsRowLabel, "totalsRowLabel must not be null");
    Objects.requireNonNull(totalsRowFunction, "totalsRowFunction must not be null");
    Objects.requireNonNull(calculatedColumnFormula, "calculatedColumnFormula must not be null");
  }

  @JsonCreator
  static TableColumnReport create(
      @JsonProperty("id") long id,
      @JsonProperty("name") String name,
      @JsonProperty("uniqueName") String uniqueName,
      @JsonProperty("totalsRowLabel") String totalsRowLabel,
      @JsonProperty("totalsRowFunction") String totalsRowFunction,
      @JsonProperty("calculatedColumnFormula") String calculatedColumnFormula) {
    return new TableColumnReport(
        id,
        name,
        uniqueName == null ? "" : uniqueName,
        totalsRowLabel == null ? "" : totalsRowLabel,
        totalsRowFunction == null ? "" : totalsRowFunction,
        calculatedColumnFormula == null ? "" : calculatedColumnFormula);
  }
}
