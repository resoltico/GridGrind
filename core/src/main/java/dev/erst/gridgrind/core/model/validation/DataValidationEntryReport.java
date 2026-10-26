package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual report for one data-validation structure read from a sheet. */
@Union(discriminator = "type")
public sealed interface DataValidationEntryReport
    permits DataValidationEntryReport.Supported, DataValidationEntryReport.Unsupported {

  /** Covered A1-style ranges stored on the sheet for this validation structure. */
  List<RangeLabel> ranges();

  /** Fully supported data-validation structure. */
  @Variant(id = "SUPPORTED", summary = "Fully supported data-validation structure.")
  record Supported(List<RangeLabel> ranges, DataValidationDefinitionReport validation)
      implements DataValidationEntryReport {
    public Supported {
      ranges = copyRanges(ranges);
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Present workbook structure that GridGrind can detect but not yet model fully. */
  @Variant(
      id = "UNSUPPORTED",
      summary = "Present workbook structure that GridGrind can detect but not yet model fully.")
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
      Optional<DataValidationPromptInput> prompt,
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
