package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Shared validation helpers for workbook-step compact constructors. */
final class WorkbookStepValidation {
  private static final Map<Class<? extends MutationAction>, Class<? extends Selector>[]>
      MUTATION_TARGET_TYPES =
          Map.ofEntries(
              Map.entry(
                  WorkbookMutationAction.EnsureSheet.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.RenameSheet.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.DeleteSheet.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.MoveSheet.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.CopySheet.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.SetActiveSheet.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.SetSelectedSheets.class,
                  targetTypes(SheetSelector.ByNames.class)),
              Map.entry(
                  WorkbookMutationAction.SetSheetVisibility.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.SetSheetProtection.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.ClearSheetProtection.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.SetWorkbookProtection.class,
                  targetTypes(WorkbookSelector.class)),
              Map.entry(
                  WorkbookMutationAction.ClearWorkbookProtection.class,
                  targetTypes(WorkbookSelector.class)),
              Map.entry(
                  WorkbookMutationAction.MergeCells.class,
                  targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  WorkbookMutationAction.UnmergeCells.class,
                  targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  WorkbookMutationAction.SetColumnWidth.class,
                  targetTypes(ColumnBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.SetRowHeight.class,
                  targetTypes(RowBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.InsertRows.class,
                  targetTypes(RowBandSelector.Insertion.class)),
              Map.entry(
                  WorkbookMutationAction.DeleteRows.class, targetTypes(RowBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.ShiftRows.class, targetTypes(RowBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.InsertColumns.class,
                  targetTypes(ColumnBandSelector.Insertion.class)),
              Map.entry(
                  WorkbookMutationAction.DeleteColumns.class,
                  targetTypes(ColumnBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.ShiftColumns.class,
                  targetTypes(ColumnBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.SetRowVisibility.class,
                  targetTypes(RowBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.SetColumnVisibility.class,
                  targetTypes(ColumnBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.GroupRows.class, targetTypes(RowBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.UngroupRows.class,
                  targetTypes(RowBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.GroupColumns.class,
                  targetTypes(ColumnBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.UngroupColumns.class,
                  targetTypes(ColumnBandSelector.Span.class)),
              Map.entry(
                  WorkbookMutationAction.SetSheetPane.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.SetSheetZoom.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.SetSheetPresentation.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.SetPrintLayout.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.ClearPrintLayout.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  CellMutationAction.SetCell.class,
                  targetTypes(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class)),
              Map.entry(
                  CellMutationAction.SetRange.class, targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  CellMutationAction.ClearRange.class, targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  CellMutationAction.SetArrayFormula.class,
                  targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  CellMutationAction.ClearArrayFormula.class,
                  targetTypes(CellSelector.ByAddress.class)),
              Map.entry(
                  StructuredMutationAction.ImportCustomXmlMapping.class,
                  targetTypes(WorkbookSelector.class)),
              Map.entry(
                  CellMutationAction.SetHyperlink.class,
                  targetTypes(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class)),
              Map.entry(
                  CellMutationAction.ClearHyperlink.class,
                  targetTypes(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class)),
              Map.entry(
                  CellMutationAction.SetComment.class,
                  targetTypes(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class)),
              Map.entry(
                  CellMutationAction.ClearComment.class,
                  targetTypes(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class)),
              Map.entry(
                  DrawingMutationAction.SetPicture.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  DrawingMutationAction.SetSignatureLine.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  DrawingMutationAction.SetChart.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  StructuredMutationAction.SetPivotTable.class,
                  targetTypes(PivotTableSelector.ByNameOnSheet.class)),
              Map.entry(
                  DrawingMutationAction.SetShape.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  DrawingMutationAction.SetEmbeddedObject.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  DrawingMutationAction.SetDrawingObjectAnchor.class,
                  targetTypes(DrawingObjectSelector.ByName.class)),
              Map.entry(
                  DrawingMutationAction.DeleteDrawingObject.class,
                  targetTypes(DrawingObjectSelector.ByName.class)),
              Map.entry(
                  CellMutationAction.ApplyStyle.class, targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  StructuredMutationAction.SetDataValidation.class,
                  targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  StructuredMutationAction.ClearDataValidations.class,
                  targetTypes(RangeSelector.class)),
              Map.entry(
                  StructuredMutationAction.SetConditionalFormatting.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  StructuredMutationAction.ClearConditionalFormatting.class,
                  targetTypes(RangeSelector.class)),
              Map.entry(
                  StructuredMutationAction.SetAutofilter.class,
                  targetTypes(RangeSelector.ByRange.class)),
              Map.entry(
                  StructuredMutationAction.ClearAutofilter.class,
                  targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  StructuredMutationAction.SetTable.class,
                  targetTypes(TableSelector.ByNameOnSheet.class)),
              Map.entry(
                  StructuredMutationAction.DeleteTable.class,
                  targetTypes(TableSelector.ByNameOnSheet.class)),
              Map.entry(
                  StructuredMutationAction.DeletePivotTable.class,
                  targetTypes(PivotTableSelector.ByNameOnSheet.class)),
              Map.entry(
                  StructuredMutationAction.SetNamedRange.class,
                  targetTypes(
                      NamedRangeSelector.ByName.class,
                      NamedRangeSelector.WorkbookScope.class,
                      NamedRangeSelector.SheetScope.class)),
              Map.entry(
                  StructuredMutationAction.DeleteNamedRange.class,
                  targetTypes(
                      NamedRangeSelector.ByName.class,
                      NamedRangeSelector.WorkbookScope.class,
                      NamedRangeSelector.SheetScope.class)),
              Map.entry(
                  CellMutationAction.AppendRow.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  WorkbookMutationAction.AutoSizeColumns.class,
                  targetTypes(SheetSelector.ByName.class)));

  private static final Map<Class<? extends InspectionQuery>, Class<? extends Selector>[]>
      INSPECTION_TARGET_TYPES =
          Map.ofEntries(
              Map.entry(
                  InspectionQuery.GetWorkbookSummary.class, targetTypes(WorkbookSelector.class)),
              Map.entry(
                  InspectionQuery.GetPackageSecurity.class, targetTypes(WorkbookSelector.class)),
              Map.entry(
                  InspectionQuery.GetWorkbookProtection.class, targetTypes(WorkbookSelector.class)),
              Map.entry(
                  InspectionQuery.GetCustomXmlMappings.class, targetTypes(WorkbookSelector.class)),
              Map.entry(
                  InspectionQuery.ExportCustomXmlMapping.class,
                  targetTypes(WorkbookSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeWorkbookFindings.class,
                  targetTypes(WorkbookSelector.class)),
              Map.entry(
                  InspectionQuery.GetNamedRanges.class, targetTypes(NamedRangeSelector.class)),
              Map.entry(
                  InspectionQuery.GetNamedRangeSurface.class,
                  targetTypes(NamedRangeSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeNamedRangeHealth.class,
                  targetTypes(NamedRangeSelector.class)),
              Map.entry(
                  InspectionQuery.GetSheetSummary.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(InspectionQuery.GetArrayFormulas.class, targetTypes(SheetSelector.class)),
              Map.entry(
                  InspectionQuery.GetMergedRegions.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  InspectionQuery.GetSheetLayout.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  InspectionQuery.GetPrintLayout.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(
                  InspectionQuery.GetAutofilters.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(InspectionQuery.GetFormulaSurface.class, targetTypes(SheetSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeFormulaHealth.class, targetTypes(SheetSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeDataValidationHealth.class,
                  targetTypes(SheetSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeConditionalFormattingHealth.class,
                  targetTypes(SheetSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeAutofilterHealth.class, targetTypes(SheetSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeHyperlinkHealth.class, targetTypes(SheetSelector.class)),
              Map.entry(
                  InspectionQuery.GetCells.class,
                  targetTypes(
                      CellSelector.ByAddress.class,
                      CellSelector.ByAddresses.class,
                      TableCellSelector.ByColumnName.class)),
              Map.entry(
                  InspectionQuery.GetHyperlinks.class,
                  targetTypes(
                      CellSelector.AllUsedInSheet.class,
                      CellSelector.ByAddress.class,
                      CellSelector.ByAddresses.class,
                      TableCellSelector.ByColumnName.class)),
              Map.entry(
                  InspectionQuery.GetComments.class,
                  targetTypes(
                      CellSelector.AllUsedInSheet.class,
                      CellSelector.ByAddress.class,
                      CellSelector.ByAddresses.class,
                      TableCellSelector.ByColumnName.class)),
              Map.entry(
                  InspectionQuery.GetWindow.class,
                  targetTypes(RangeSelector.RectangularWindow.class)),
              Map.entry(
                  InspectionQuery.GetSheetSchema.class,
                  targetTypes(RangeSelector.RectangularWindow.class)),
              Map.entry(InspectionQuery.GetDataValidations.class, targetTypes(RangeSelector.class)),
              Map.entry(
                  InspectionQuery.GetConditionalFormatting.class, targetTypes(RangeSelector.class)),
              Map.entry(
                  InspectionQuery.GetDrawingObjects.class,
                  targetTypes(DrawingObjectSelector.AllOnSheet.class)),
              Map.entry(
                  InspectionQuery.GetDrawingObjectPayload.class,
                  targetTypes(DrawingObjectSelector.ByName.class)),
              Map.entry(
                  InspectionQuery.GetCharts.class,
                  targetTypes(ChartSelector.AllOnSheet.class, SheetSelector.ByName.class)),
              Map.entry(
                  InspectionQuery.GetPivotTables.class, targetTypes(PivotTableSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzePivotTableHealth.class,
                  targetTypes(PivotTableSelector.class)),
              Map.entry(InspectionQuery.GetTables.class, targetTypes(TableSelector.class)),
              Map.entry(
                  InspectionQuery.AnalyzeTableHealth.class, targetTypes(TableSelector.class)));

  private static final Map<Class<? extends Assertion>, Class<? extends Selector>[]>
      ASSERTION_TARGET_TYPES =
          Map.ofEntries(
              Map.entry(Assertion.NamedRangePresent.class, targetTypes(NamedRangeSelector.class)),
              Map.entry(Assertion.NamedRangeAbsent.class, targetTypes(NamedRangeSelector.class)),
              Map.entry(Assertion.TablePresent.class, targetTypes(TableSelector.class)),
              Map.entry(Assertion.TableAbsent.class, targetTypes(TableSelector.class)),
              Map.entry(Assertion.PivotTablePresent.class, targetTypes(PivotTableSelector.class)),
              Map.entry(Assertion.PivotTableAbsent.class, targetTypes(PivotTableSelector.class)),
              Map.entry(Assertion.ChartPresent.class, targetTypes(ChartSelector.class)),
              Map.entry(Assertion.ChartAbsent.class, targetTypes(ChartSelector.class)),
              Map.entry(
                  Assertion.CellValue.class,
                  targetTypes(
                      CellSelector.ByAddress.class,
                      CellSelector.ByAddresses.class,
                      TableCellSelector.ByColumnName.class)),
              Map.entry(
                  Assertion.DisplayValue.class,
                  targetTypes(
                      CellSelector.ByAddress.class,
                      CellSelector.ByAddresses.class,
                      TableCellSelector.ByColumnName.class)),
              Map.entry(
                  Assertion.FormulaText.class,
                  targetTypes(
                      CellSelector.ByAddress.class,
                      CellSelector.ByAddresses.class,
                      TableCellSelector.ByColumnName.class)),
              Map.entry(
                  Assertion.CellStyle.class,
                  targetTypes(
                      CellSelector.ByAddress.class,
                      CellSelector.ByAddresses.class,
                      TableCellSelector.ByColumnName.class)),
              Map.entry(
                  Assertion.WorkbookProtectionFacts.class, targetTypes(WorkbookSelector.class)),
              Map.entry(
                  Assertion.SheetStructureFacts.class, targetTypes(SheetSelector.ByName.class)),
              Map.entry(Assertion.NamedRangeFacts.class, targetTypes(NamedRangeSelector.class)),
              Map.entry(Assertion.TableFacts.class, targetTypes(TableSelector.class)),
              Map.entry(Assertion.PivotTableFacts.class, targetTypes(PivotTableSelector.class)),
              Map.entry(Assertion.ChartFacts.class, targetTypes(ChartSelector.class)),
              Map.entry(Assertion.AllOf.class, targetTypes(Selector.class)),
              Map.entry(Assertion.AnyOf.class, targetTypes(Selector.class)),
              Map.entry(Assertion.Not.class, targetTypes(Selector.class)));

  private WorkbookStepValidation() {}

  static String requireStepId(String stepId) {
    Objects.requireNonNull(stepId, "stepId must not be null");
    if (stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    return stepId;
  }

  static Selector requireTarget(Selector target) {
    Objects.requireNonNull(target, "target must not be null");
    return target;
  }

  static MutationAction requireCompatible(Selector target, MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    requireTargetType(target, action.actionType(), allowedTargetTypes(action));
    return action;
  }

  static InspectionQuery requireCompatible(Selector target, InspectionQuery query) {
    Objects.requireNonNull(query, "query must not be null");
    requireTargetType(target, query.queryType(), allowedTargetTypes(query));
    return query;
  }

  static Assertion requireCompatible(Selector target, Assertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    requireTargetType(target, assertion.assertionType(), allowedTargetTypes(assertion));
    return assertion;
  }

  static Class<? extends Selector>[] allowedTargetTypes(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    return allowedTargetTypesForMutationActionType(action.getClass());
  }

  static Class<? extends Selector>[] allowedTargetTypes(InspectionQuery query) {
    Objects.requireNonNull(query, "query must not be null");
    return allowedTargetTypesForInspectionQueryType(query.getClass());
  }

  static Class<? extends Selector>[] allowedTargetTypes(Assertion assertion) {
    return AssertionTargetingSupport.allowedTargetTypes(assertion, ASSERTION_TARGET_TYPES);
  }

  static Class<? extends Selector>[] allowedTargetTypesForMutationActionType(
      Class<? extends MutationAction> actionType) {
    Objects.requireNonNull(actionType, "actionType must not be null");
    return configuredTargetTypes(actionType, MUTATION_TARGET_TYPES, "action");
  }

  static Class<? extends Selector>[] allowedTargetTypesForInspectionQueryType(
      Class<? extends InspectionQuery> queryType) {
    Objects.requireNonNull(queryType, "queryType must not be null");
    return configuredTargetTypes(queryType, INSPECTION_TARGET_TYPES, "query");
  }

  static Class<? extends Selector>[] staticAllowedTargetTypesForAssertionType(
      Class<? extends Assertion> assertionType) {
    return AssertionTargetingSupport.staticAllowedTargetTypesForAssertionType(
        assertionType, ASSERTION_TARGET_TYPES);
  }

  static Optional<String> dynamicTargetSelectorRuleForAssertionType(
      Class<? extends Assertion> assertionType) {
    return AssertionTargetingSupport.dynamicTargetSelectorRuleForAssertionType(assertionType);
  }

  static Optional<String> targetSelectorRuleForAssertionType(
      Class<? extends Assertion> assertionType) {
    return AssertionTargetingSupport.targetSelectorRuleForAssertionType(assertionType);
  }

  private static <T> Class<? extends Selector>[] configuredTargetTypes(
      Class<?> stepClass,
      Map<Class<? extends T>, Class<? extends Selector>[]> mappings,
      String mappingLabel) {
    return Objects.requireNonNull(
        mappings.get(stepClass),
        "No target-type mapping configured for " + mappingLabel + " class " + stepClass.getName());
  }

  @SafeVarargs
  private static void requireTargetType(
      Selector target, String stepType, Class<? extends Selector>... allowedTypes) {
    for (Class<? extends Selector> allowedType : allowedTypes) {
      if (allowedType.isInstance(target)) {
        return;
      }
    }
    throw new IllegalArgumentException(
        stepType
            + " requires target type "
            + humanTargetTypes(allowedTypes)
            + " but got "
            + target.getClass().getSimpleName());
  }

  @SafeVarargs
  private static Class<? extends Selector>[] targetTypes(Class<? extends Selector>... types) {
    return types;
  }

  private static String humanTargetTypes(Class<? extends Selector>[] allowedTypes) {
    if (allowedTypes.length == 1) {
      return allowedTypes[0].getSimpleName();
    }
    StringBuilder builder = new StringBuilder();
    for (int index = 0; index < allowedTypes.length; index++) {
      if (index > 0) {
        builder.append(index == allowedTypes.length - 1 ? " or " : ", ");
      }
      builder.append(allowedTypes[index].getSimpleName());
    }
    return builder.toString();
  }
}
