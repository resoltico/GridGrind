package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationSnapshot;
import java.util.List;

/** Validation, formatting, autofilter, and table inventory results. */
public sealed interface WorkbookRuleResult extends WorkbookReadIntrospectionResult
    permits WorkbookRuleResult.DataValidationsResult,
        WorkbookRuleResult.ConditionalFormattingResult,
        WorkbookRuleResult.AutofiltersResult,
        WorkbookRuleResult.TablesResult {

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  record DataValidationsResult(
      StepId stepId, String sheetName, List<ExcelDataValidationSnapshot> validations)
      implements WorkbookRuleResult {
    public DataValidationsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      validations = copyValues(validations, "validations");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record ConditionalFormattingResult(
      StepId stepId,
      String sheetName,
      List<ExcelConditionalFormattingBlockSnapshot> conditionalFormattingBlocks)
      implements WorkbookRuleResult {
    public ConditionalFormattingResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      conditionalFormattingBlocks =
          copyValues(conditionalFormattingBlocks, "conditionalFormattingBlocks");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  record AutofiltersResult(
      StepId stepId, String sheetName, List<ExcelAutofilterSnapshot> autofilters)
      implements WorkbookRuleResult {
    public AutofiltersResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      autofilters = copyValues(autofilters, "autofilters");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  record TablesResult(StepId stepId, List<ExcelTableSnapshot> tables)
      implements WorkbookRuleResult {
    public TablesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      tables = copyValues(tables, "tables");
    }
  }
}
