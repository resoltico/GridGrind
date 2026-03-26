package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindResponse;
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

  /** Counts analysis sheet inspections requested by the protocol workflow generator. */
  public static int analysisSheetCount(GridGrindRequest request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.analysis().sheets().size();
  }

  /** Returns the stable workbook source-mode label for one request. */
  public static String sourceKind(GridGrindRequest request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.New _ -> "NEW";
      case GridGrindRequest.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  /** Returns the stable workbook persistence-mode label for one request. */
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
      case WorkbookOperation.ApplyStyle _ -> "APPLY_STYLE";
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
      case WorkbookCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
      case WorkbookCommand.EvaluateAllFormulas _ -> "EVALUATE_ALL_FORMULAS";
      case WorkbookCommand.ForceFormulaRecalculationOnOpen _ ->
          "FORCE_FORMULA_RECALCULATION_ON_OPEN";
    };
  }

  private static void increment(Map<String, Long> counts, String key) {
    counts.merge(key, 1L, Long::sum);
  }
}
