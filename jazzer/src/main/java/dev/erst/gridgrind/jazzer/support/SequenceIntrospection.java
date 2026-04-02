package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.WorkbookOperation;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Summarizes generated workflow and command shapes for reports, replay, and telemetry. */
public final class SequenceIntrospection {
  private SequenceIntrospection() {}

  /** Counts protocol operation kinds by stable wire-facing name. */
  public static Map<String, Long> operationKinds(List<WorkbookOperation> operations) {
    Objects.requireNonNull(operations, "operations must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    operations.forEach(operation -> increment(counts, operationKind(operation)));
    return Map.copyOf(counts);
  }

  /** Counts workbook command kinds by stable engine-facing name. */
  public static Map<String, Long> commandKinds(List<WorkbookCommand> commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    commands.forEach(command -> increment(counts, commandKind(command)));
    return Map.copyOf(counts);
  }

  /** Counts workbook read kinds by stable wire-facing name. */
  public static Map<String, Long> readKinds(List<WorkbookReadOperation> reads) {
    Objects.requireNonNull(reads, "reads must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    reads.forEach(read -> increment(counts, readKind(read)));
    return Map.copyOf(counts);
  }

  /** Counts style-attribute labels observed in protocol style operations. */
  public static Map<String, Long> styleKinds(List<WorkbookOperation> operations) {
    Objects.requireNonNull(operations, "operations must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    operations.forEach(
        operation -> {
          if (operation instanceof WorkbookOperation.ApplyStyle applyStyle) {
            StyleKindIntrospection.styleKinds(applyStyle.style())
                .forEach((key, value) -> counts.merge(key, value, Long::sum));
          }
        });
    return Map.copyOf(counts);
  }

  /** Counts style-attribute labels observed in engine style commands. */
  public static Map<String, Long> styleKindsFromCommands(List<WorkbookCommand> commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    commands.forEach(
        command -> {
          if (command instanceof WorkbookCommand.ApplyStyle applyStyle) {
            StyleKindIntrospection.styleKinds(applyStyle.style())
                .forEach((key, value) -> counts.merge(key, value, Long::sum));
          }
        });
    return Map.copyOf(counts);
  }

  /** Returns the response family label used in telemetry and summaries. */
  public static String responseKind(GridGrindResponse response) {
    Objects.requireNonNull(response, "response must not be null");
    return switch (response) {
      case GridGrindResponse.Success _ -> "SUCCESS";
      case GridGrindResponse.Failure _ -> "FAILURE";
    };
  }

  /** Counts ordered workbook reads requested by the protocol workflow generator. */
  public static int readCount(GridGrindRequest request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.reads().size();
  }

  /** Returns the stable workbook source-type label for one request. */
  public static String sourceKind(GridGrindRequest request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.New _ -> "NEW";
      case GridGrindRequest.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  /** Returns the stable workbook persistence-type label for one request. */
  public static String persistenceKind(GridGrindRequest request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.persistence()) {
      case GridGrindRequest.WorkbookPersistence.None _ -> "NONE";
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
      case GridGrindRequest.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  /** Returns the stable wire-facing name for one protocol operation variant. */
  static String operationKind(WorkbookOperation operation) {
    Objects.requireNonNull(operation, "operation must not be null");
    return switch (operation) {
      case WorkbookOperation.EnsureSheet _ -> "ENSURE_SHEET";
      case WorkbookOperation.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookOperation.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookOperation.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookOperation.MergeCells _ -> "MERGE_CELLS";
      case WorkbookOperation.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookOperation.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookOperation.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookOperation.FreezePanes _ -> "FREEZE_PANES";
      case WorkbookOperation.SetCell _ -> "SET_CELL";
      case WorkbookOperation.SetRange _ -> "SET_RANGE";
      case WorkbookOperation.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookOperation.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookOperation.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookOperation.SetComment _ -> "SET_COMMENT";
      case WorkbookOperation.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookOperation.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookOperation.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookOperation.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookOperation.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookOperation.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookOperation.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookOperation.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookOperation.SetTable _ -> "SET_TABLE";
      case WorkbookOperation.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookOperation.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookOperation.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookOperation.AppendRow _ -> "APPEND_ROW";
      case WorkbookOperation.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case WorkbookOperation.EvaluateFormulas _ -> "EVALUATE_FORMULAS";
      case WorkbookOperation.ForceFormulaRecalculationOnOpen _ ->
          "FORCE_FORMULA_RECALCULATION_ON_OPEN";
    };
  }

  /** Returns the stable engine-facing name for one workbook command variant. */
  static String commandKind(WorkbookCommand command) {
    Objects.requireNonNull(command, "command must not be null");
    return switch (command) {
      case WorkbookCommand.CreateSheet _ -> "CREATE_SHEET";
      case WorkbookCommand.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookCommand.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookCommand.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookCommand.MergeCells _ -> "MERGE_CELLS";
      case WorkbookCommand.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookCommand.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookCommand.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookCommand.FreezePanes _ -> "FREEZE_PANES";
      case WorkbookCommand.SetCell _ -> "SET_CELL";
      case WorkbookCommand.SetRange _ -> "SET_RANGE";
      case WorkbookCommand.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookCommand.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookCommand.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookCommand.SetComment _ -> "SET_COMMENT";
      case WorkbookCommand.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookCommand.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookCommand.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookCommand.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookCommand.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookCommand.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookCommand.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookCommand.SetTable _ -> "SET_TABLE";
      case WorkbookCommand.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookCommand.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookCommand.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case WorkbookCommand.EvaluateAllFormulas _ -> "EVALUATE_ALL_FORMULAS";
      case WorkbookCommand.ForceFormulaRecalculationOnOpen _ ->
          "FORCE_FORMULA_RECALCULATION_ON_OPEN";
    };
  }

  /** Returns the stable wire-facing name for one workbook read variant. */
  static String readKind(WorkbookReadOperation read) {
    Objects.requireNonNull(read, "read must not be null");
    return switch (read) {
      case WorkbookReadOperation.GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadOperation.GetNamedRanges _ -> "GET_NAMED_RANGES";
      case WorkbookReadOperation.GetSheetSummary _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadOperation.GetCells _ -> "GET_CELLS";
      case WorkbookReadOperation.GetWindow _ -> "GET_WINDOW";
      case WorkbookReadOperation.GetMergedRegions _ -> "GET_MERGED_REGIONS";
      case WorkbookReadOperation.GetHyperlinks _ -> "GET_HYPERLINKS";
      case WorkbookReadOperation.GetComments _ -> "GET_COMMENTS";
      case WorkbookReadOperation.GetSheetLayout _ -> "GET_SHEET_LAYOUT";
      case WorkbookReadOperation.GetDataValidations _ -> "GET_DATA_VALIDATIONS";
      case WorkbookReadOperation.GetConditionalFormatting _ -> "GET_CONDITIONAL_FORMATTING";
      case WorkbookReadOperation.GetAutofilters _ -> "GET_AUTOFILTERS";
      case WorkbookReadOperation.GetTables _ -> "GET_TABLES";
      case WorkbookReadOperation.GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadOperation.GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadOperation.GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
      case WorkbookReadOperation.AnalyzeFormulaHealth _ -> "ANALYZE_FORMULA_HEALTH";
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ -> "ANALYZE_DATA_VALIDATION_HEALTH";
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ ->
          "ANALYZE_CONDITIONAL_FORMATTING_HEALTH";
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ -> "ANALYZE_AUTOFILTER_HEALTH";
      case WorkbookReadOperation.AnalyzeTableHealth _ -> "ANALYZE_TABLE_HEALTH";
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ -> "ANALYZE_HYPERLINK_HEALTH";
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ -> "ANALYZE_NAMED_RANGE_HEALTH";
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> "ANALYZE_WORKBOOK_FINDINGS";
    };
  }

  private static void increment(Map<String, Long> counts, String key) {
    counts.merge(key, 1L, Long::sum);
  }
}
