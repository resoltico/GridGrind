package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Summarizes generated workflow and command shapes for reports, replay, and telemetry. */
public final class SequenceIntrospection {
  private SequenceIntrospection() {}

  /** Counts protocol mutation kinds by stable wire-facing action name. */
  public static Map<String, Long> mutationKinds(List<MutationStep> mutations) {
    Objects.requireNonNull(mutations, "mutations must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    mutations.forEach(mutation -> increment(counts, mutationKind(mutation)));
    return Map.copyOf(counts);
  }

  /** Counts protocol inspection kinds by stable wire-facing query name. */
  public static Map<String, Long> inspectionKinds(List<InspectionStep> inspections) {
    Objects.requireNonNull(inspections, "inspections must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    inspections.forEach(inspection -> increment(counts, inspectionKind(inspection)));
    return Map.copyOf(counts);
  }

  /** Counts protocol assertion kinds by stable wire-facing assertion name. */
  public static Map<String, Long> assertionKinds(List<AssertionStep> assertions) {
    Objects.requireNonNull(assertions, "assertions must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    assertions.forEach(assertion -> increment(counts, assertionKind(assertion)));
    return Map.copyOf(counts);
  }

  /** Counts workbook command kinds by stable engine-facing name. */
  public static Map<String, Long> commandKinds(List<WorkbookCommand> commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    commands.forEach(command -> increment(counts, commandKind(command)));
    return Map.copyOf(counts);
  }

  /** Counts style-attribute labels observed in protocol style mutations. */
  public static Map<String, Long> styleKinds(List<MutationStep> mutations) {
    Objects.requireNonNull(mutations, "mutations must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    mutations.forEach(
        mutation -> {
          if (mutation.action()
              instanceof dev.erst.gridgrind.contract.action.MutationAction.ApplyStyle applyStyle) {
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

  /** Counts ordered inspection steps requested by one authored workflow. */
  public static int inspectionCount(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.inspectionSteps().size();
  }

  /** Counts ordered assertion steps requested by one authored workflow. */
  public static int assertionCount(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.assertionSteps().size();
  }

  /** Returns the stable workbook source-type label for one request. */
  public static String sourceKind(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> "NEW";
      case WorkbookPlan.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  /** Returns the stable workbook persistence-type label for one request. */
  public static String persistenceKind(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  /** Returns the stable wire-facing name for one protocol mutation step. */
  static String mutationKind(MutationStep mutation) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    return mutation.action().actionType();
  }

  /** Returns the stable engine-facing name for one workbook command variant. */
  static String commandKind(WorkbookCommand command) {
    Objects.requireNonNull(command, "command must not be null");
    return switch (command) {
      case WorkbookCommand.CreateSheet _ -> "CREATE_SHEET";
      case WorkbookCommand.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookCommand.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookCommand.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookCommand.CopySheet _ -> "COPY_SHEET";
      case WorkbookCommand.SetActiveSheet _ -> "SET_ACTIVE_SHEET";
      case WorkbookCommand.SetSelectedSheets _ -> "SET_SELECTED_SHEETS";
      case WorkbookCommand.SetSheetVisibility _ -> "SET_SHEET_VISIBILITY";
      case WorkbookCommand.SetSheetProtection _ -> "SET_SHEET_PROTECTION";
      case WorkbookCommand.ClearSheetProtection _ -> "CLEAR_SHEET_PROTECTION";
      case WorkbookCommand.SetWorkbookProtection _ -> "SET_WORKBOOK_PROTECTION";
      case WorkbookCommand.ClearWorkbookProtection _ -> "CLEAR_WORKBOOK_PROTECTION";
      case WorkbookCommand.MergeCells _ -> "MERGE_CELLS";
      case WorkbookCommand.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookCommand.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookCommand.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookCommand.InsertRows _ -> "INSERT_ROWS";
      case WorkbookCommand.DeleteRows _ -> "DELETE_ROWS";
      case WorkbookCommand.ShiftRows _ -> "SHIFT_ROWS";
      case WorkbookCommand.InsertColumns _ -> "INSERT_COLUMNS";
      case WorkbookCommand.DeleteColumns _ -> "DELETE_COLUMNS";
      case WorkbookCommand.ShiftColumns _ -> "SHIFT_COLUMNS";
      case WorkbookCommand.SetRowVisibility _ -> "SET_ROW_VISIBILITY";
      case WorkbookCommand.SetColumnVisibility _ -> "SET_COLUMN_VISIBILITY";
      case WorkbookCommand.GroupRows _ -> "GROUP_ROWS";
      case WorkbookCommand.UngroupRows _ -> "UNGROUP_ROWS";
      case WorkbookCommand.GroupColumns _ -> "GROUP_COLUMNS";
      case WorkbookCommand.UngroupColumns _ -> "UNGROUP_COLUMNS";
      case WorkbookCommand.SetSheetPane _ -> "SET_SHEET_PANE";
      case WorkbookCommand.SetSheetZoom _ -> "SET_SHEET_ZOOM";
      case WorkbookCommand.SetSheetPresentation _ -> "SET_SHEET_PRESENTATION";
      case WorkbookCommand.SetPrintLayout _ -> "SET_PRINT_LAYOUT";
      case WorkbookCommand.ClearPrintLayout _ -> "CLEAR_PRINT_LAYOUT";
      case WorkbookCommand.SetCell _ -> "SET_CELL";
      case WorkbookCommand.SetRange _ -> "SET_RANGE";
      case WorkbookCommand.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookCommand.SetArrayFormula _ -> "SET_ARRAY_FORMULA";
      case WorkbookCommand.ClearArrayFormula _ -> "CLEAR_ARRAY_FORMULA";
      case WorkbookCommand.ImportCustomXmlMapping _ -> "IMPORT_CUSTOM_XML_MAPPING";
      case WorkbookCommand.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookCommand.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookCommand.SetComment _ -> "SET_COMMENT";
      case WorkbookCommand.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookCommand.SetPicture _ -> "SET_PICTURE";
      case WorkbookCommand.SetSignatureLine _ -> "SET_SIGNATURE_LINE";
      case WorkbookCommand.SetChart _ -> "SET_CHART";
      case WorkbookCommand.SetPivotTable _ -> "SET_PIVOT_TABLE";
      case WorkbookCommand.SetShape _ -> "SET_SHAPE";
      case WorkbookCommand.SetEmbeddedObject _ -> "SET_EMBEDDED_OBJECT";
      case WorkbookCommand.SetDrawingObjectAnchor _ -> "SET_DRAWING_OBJECT_ANCHOR";
      case WorkbookCommand.DeleteDrawingObject _ -> "DELETE_DRAWING_OBJECT";
      case WorkbookCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookCommand.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookCommand.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookCommand.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookCommand.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookCommand.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookCommand.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookCommand.SetTable _ -> "SET_TABLE";
      case WorkbookCommand.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookCommand.DeletePivotTable _ -> "DELETE_PIVOT_TABLE";
      case WorkbookCommand.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookCommand.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
    };
  }

  /** Returns the stable wire-facing name for one protocol inspection step. */
  static String inspectionKind(InspectionStep inspection) {
    Objects.requireNonNull(inspection, "inspection must not be null");
    return inspection.query().queryType();
  }

  /** Returns the stable wire-facing name for one protocol assertion step. */
  static String assertionKind(AssertionStep assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return assertion.assertion().assertionType();
  }

  private static void increment(Map<String, Long> counts, String key) {
    counts.merge(key, 1L, Long::sum);
  }
}
