package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import java.util.List;

/** Validation, formatting, autofilter, and table inventory results. */
public sealed interface WorkbookRuleResult extends WorkbookReadIntrospectionResult
    permits WorkbookRuleResult.DataValidationsResult,
        WorkbookRuleResult.ConditionalFormattingResult,
        WorkbookRuleResult.AutofiltersResult,
        WorkbookRuleResult.TablesResult {

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  record DataValidationsResult(
      String stepId, String sheetName, List<ExcelDataValidationSnapshot> validations)
      implements WorkbookRuleResult {
    public DataValidationsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      validations = copyValues(validations, "validations");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record ConditionalFormattingResult(
      String stepId,
      String sheetName,
      List<ExcelConditionalFormattingBlockSnapshot> conditionalFormattingBlocks)
      implements WorkbookRuleResult {
    public ConditionalFormattingResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      conditionalFormattingBlocks =
          copyValues(conditionalFormattingBlocks, "conditionalFormattingBlocks");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  record AutofiltersResult(
      String stepId, String sheetName, List<ExcelAutofilterSnapshot> autofilters)
      implements WorkbookRuleResult {
    public AutofiltersResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      autofilters = copyValues(autofilters, "autofilters");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  record TablesResult(String stepId, List<ExcelTableSnapshot> tables)
      implements WorkbookRuleResult {
    public TablesResult {
      stepId = requireNonBlank(stepId, "stepId");
      tables = copyValues(tables, "tables");
    }
  }
}
