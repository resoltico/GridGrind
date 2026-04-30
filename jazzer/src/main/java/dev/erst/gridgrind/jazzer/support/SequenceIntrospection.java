package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.*;
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
              instanceof
              dev.erst.gridgrind.contract.action.CellMutationAction.ApplyStyle applyStyle) {
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
          if (command instanceof WorkbookFormattingCommand.ApplyStyle applyStyle) {
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
    return request.stepPartition().inspections().size();
  }

  /** Counts ordered assertion steps requested by one authored workflow. */
  public static int assertionCount(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.stepPartition().assertions().size();
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
      case WorkbookSheetCommand.CreateSheet _ -> "CREATE_SHEET";
      case WorkbookSheetCommand.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookSheetCommand.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookSheetCommand.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookSheetCommand.CopySheet _ -> "COPY_SHEET";
      case WorkbookSheetCommand.SetActiveSheet _ -> "SET_ACTIVE_SHEET";
      case WorkbookSheetCommand.SetSelectedSheets _ -> "SET_SELECTED_SHEETS";
      case WorkbookSheetCommand.SetSheetVisibility _ -> "SET_SHEET_VISIBILITY";
      case WorkbookSheetCommand.SetSheetProtection _ -> "SET_SHEET_PROTECTION";
      case WorkbookSheetCommand.ClearSheetProtection _ -> "CLEAR_SHEET_PROTECTION";
      case WorkbookSheetCommand.SetWorkbookProtection _ -> "SET_WORKBOOK_PROTECTION";
      case WorkbookSheetCommand.ClearWorkbookProtection _ -> "CLEAR_WORKBOOK_PROTECTION";
      case WorkbookStructureCommand.MergeCells _ -> "MERGE_CELLS";
      case WorkbookStructureCommand.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookStructureCommand.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookStructureCommand.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookStructureCommand.InsertRows _ -> "INSERT_ROWS";
      case WorkbookStructureCommand.DeleteRows _ -> "DELETE_ROWS";
      case WorkbookStructureCommand.ShiftRows _ -> "SHIFT_ROWS";
      case WorkbookStructureCommand.InsertColumns _ -> "INSERT_COLUMNS";
      case WorkbookStructureCommand.DeleteColumns _ -> "DELETE_COLUMNS";
      case WorkbookStructureCommand.ShiftColumns _ -> "SHIFT_COLUMNS";
      case WorkbookStructureCommand.SetRowVisibility _ -> "SET_ROW_VISIBILITY";
      case WorkbookStructureCommand.SetColumnVisibility _ -> "SET_COLUMN_VISIBILITY";
      case WorkbookStructureCommand.GroupRows _ -> "GROUP_ROWS";
      case WorkbookStructureCommand.UngroupRows _ -> "UNGROUP_ROWS";
      case WorkbookStructureCommand.GroupColumns _ -> "GROUP_COLUMNS";
      case WorkbookStructureCommand.UngroupColumns _ -> "UNGROUP_COLUMNS";
      case WorkbookLayoutCommand.SetSheetPane _ -> "SET_SHEET_PANE";
      case WorkbookLayoutCommand.SetSheetZoom _ -> "SET_SHEET_ZOOM";
      case WorkbookLayoutCommand.SetSheetPresentation _ -> "SET_SHEET_PRESENTATION";
      case WorkbookLayoutCommand.SetPrintLayout _ -> "SET_PRINT_LAYOUT";
      case WorkbookLayoutCommand.ClearPrintLayout _ -> "CLEAR_PRINT_LAYOUT";
      case WorkbookCellCommand.SetCell _ -> "SET_CELL";
      case WorkbookCellCommand.SetRange _ -> "SET_RANGE";
      case WorkbookCellCommand.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookCellCommand.SetArrayFormula _ -> "SET_ARRAY_FORMULA";
      case WorkbookCellCommand.ClearArrayFormula _ -> "CLEAR_ARRAY_FORMULA";
      case WorkbookMetadataCommand.ImportCustomXmlMapping _ -> "IMPORT_CUSTOM_XML_MAPPING";
      case WorkbookAnnotationCommand.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookAnnotationCommand.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookAnnotationCommand.SetComment _ -> "SET_COMMENT";
      case WorkbookAnnotationCommand.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookDrawingCommand.SetPicture _ -> "SET_PICTURE";
      case WorkbookDrawingCommand.SetSignatureLine _ -> "SET_SIGNATURE_LINE";
      case WorkbookDrawingCommand.SetChart _ -> "SET_CHART";
      case WorkbookTabularCommand.SetPivotTable _ -> "SET_PIVOT_TABLE";
      case WorkbookDrawingCommand.SetShape _ -> "SET_SHAPE";
      case WorkbookDrawingCommand.SetEmbeddedObject _ -> "SET_EMBEDDED_OBJECT";
      case WorkbookDrawingCommand.SetDrawingObjectAnchor _ -> "SET_DRAWING_OBJECT_ANCHOR";
      case WorkbookDrawingCommand.DeleteDrawingObject _ -> "DELETE_DRAWING_OBJECT";
      case WorkbookFormattingCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookFormattingCommand.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookFormattingCommand.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookFormattingCommand.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookFormattingCommand.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookTabularCommand.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookTabularCommand.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookTabularCommand.SetTable _ -> "SET_TABLE";
      case WorkbookTabularCommand.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookTabularCommand.DeletePivotTable _ -> "DELETE_PIVOT_TABLE";
      case WorkbookMetadataCommand.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookMetadataCommand.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookCellCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookLayoutCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
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
