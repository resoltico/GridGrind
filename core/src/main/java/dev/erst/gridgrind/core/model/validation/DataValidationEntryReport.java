package dev.erst.gridgrind.core.model.validation;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual report for one data-validation structure read from a sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DataValidationEntryReport.Supported.class, name = "SUPPORTED"),
  @JsonSubTypes.Type(value = DataValidationEntryReport.Unsupported.class, name = "UNSUPPORTED")
})
public sealed interface DataValidationEntryReport
    permits DataValidationEntryReport.Supported, DataValidationEntryReport.Unsupported {

  /** Covered A1-style ranges stored on the sheet for this validation structure. */
  List<RangeLabel> ranges();

  /** Fully supported data-validation structure. */
  record Supported(List<RangeLabel> ranges, DataValidationDefinitionReport validation)
      implements DataValidationEntryReport {
    public Supported {
      ranges = copyRanges(ranges);
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Present workbook structure that GridGrind can detect but not yet model fully. */
  record Unsupported(List<RangeLabel> ranges, String kind, String detail)
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
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DataValidationPromptInput> prompt,
      @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<DataValidationErrorAlertInput> errorAlert) {
    public DataValidationDefinitionReport {
      Objects.requireNonNull(rule, "rule must not be null");
      Objects.requireNonNull(prompt, "prompt must not be null");
      Objects.requireNonNull(errorAlert, "errorAlert must not be null");
    }
  }

  private static List<RangeLabel> copyRanges(List<RangeLabel> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<RangeLabel> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
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
