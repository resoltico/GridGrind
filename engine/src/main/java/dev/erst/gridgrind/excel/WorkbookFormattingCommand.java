package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Formatting-rule commands for styles, validations, and conditional formatting. */
public sealed interface WorkbookFormattingCommand extends WorkbookCommand
    permits WorkbookFormattingCommand.ApplyStyle,
        WorkbookFormattingCommand.SetDataValidation,
        WorkbookFormattingCommand.ClearDataValidations,
        WorkbookFormattingCommand.SetConditionalFormatting,
        WorkbookFormattingCommand.ClearConditionalFormatting {

  record ApplyStyle(String sheetName, String range, ExcelCellStyle style)
      implements WorkbookFormattingCommand {
    public ApplyStyle {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(style, "style must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  record SetDataValidation(String sheetName, String range, ExcelDataValidationDefinition validation)
      implements WorkbookFormattingCommand {
    public SetDataValidation {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(validation, "validation must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  record ClearDataValidations(String sheetName, ExcelRangeSelection selection)
      implements WorkbookFormattingCommand {
    public ClearDataValidations {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one logical conditional-formatting block over the requested ranges. */
  record SetConditionalFormatting(String sheetName, ExcelConditionalFormattingBlockDefinition block)
      implements WorkbookFormattingCommand {
    public SetConditionalFormatting {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(block, "block must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Removes conditional-formatting blocks on the sheet that intersect the provided selection. */
  record ClearConditionalFormatting(String sheetName, ExcelRangeSelection selection)
      implements WorkbookFormattingCommand {
    public ClearConditionalFormatting {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }
}
