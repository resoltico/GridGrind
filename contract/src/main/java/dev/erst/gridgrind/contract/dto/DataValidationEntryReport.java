package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;

/** Protocol-facing factual report for one data-validation structure read from a sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DataValidationEntryReport.Supported.class, name = "SUPPORTED"),
  @JsonSubTypes.Type(value = DataValidationEntryReport.Unsupported.class, name = "UNSUPPORTED")
})
public sealed interface DataValidationEntryReport
    permits DataValidationEntryReport.Supported, DataValidationEntryReport.Unsupported {

  /** Covered A1-style ranges stored on the sheet for this validation structure. */
  List<String> ranges();

  /** Fully supported data-validation structure. */
  record Supported(List<String> ranges, DataValidationDefinitionReport validation)
      implements DataValidationEntryReport {
    public Supported {
      ranges = copyRanges(ranges);
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Present workbook structure that GridGrind can detect but not yet model fully. */
  record Unsupported(List<String> ranges, String kind, String detail)
      implements DataValidationEntryReport {
    public Unsupported {
      ranges = copyRanges(ranges);
      kind = requireNonBlank(kind, "kind");
      detail = requireNonBlank(detail, "detail");
    }
  }

  /** Protocol-facing supported validation definition. */
  record DataValidationDefinitionReport(
      DataValidationRuleInput rule,
      boolean allowBlank,
      boolean suppressDropDownArrow,
      DataValidationPromptInput prompt,
      DataValidationErrorAlertInput errorAlert) {
    public DataValidationDefinitionReport {
      Objects.requireNonNull(rule, "rule must not be null");
    }
  }

  private static List<String> copyRanges(List<String> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<String> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
    }
    for (String range : copy) {
      requireNonBlank(range, "ranges");
    }
    return copy;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
