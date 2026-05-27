package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadResult;
import java.util.List;
import java.util.Objects;

/** Shared assertion observation adapter over selector resolution and workbook read execution. */
final class AssertionObservationExecutor {
  private final WorkbookExecutionEngine workbookEngine;
  private final SemanticSelectorResolver selectorResolver;

  AssertionObservationExecutor(
      WorkbookExecutionEngine workbookEngine, SemanticSelectorResolver selectorResolver) {
    this.workbookEngine = Objects.requireNonNull(workbookEngine, "workbookEngine must not be null");
    this.selectorResolver =
        Objects.requireNonNull(selectorResolver, "selectorResolver must not be null");
  }

  InspectionResult executeObservation(
      String stepId,
      Selector target,
      InspectionQuery query,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SemanticSelectorResolver.ResolvedInspectionTarget resolvedTarget =
        selectorResolver.resolveInspectionTarget(stepId, workbook, target, query);
    if (resolvedTarget.isShortCircuit()) {
      return Objects.requireNonNull(
          resolvedTarget.shortCircuitResult(), "short-circuit observation result must be present");
    }
    WorkbookReadResult result =
        workbookEngine
            .read(
                workbook,
                workbookLocation,
                InspectionCommandConverter.toReadCommand(
                    stepId,
                    Objects.requireNonNull(
                        resolvedTarget.selector(), "resolved selector must be present"),
                    query))
            .getFirst();
    return InspectionResultConverter.toReadResult(result);
  }

  InspectionResult presenceObservation(
      String stepId, Selector target, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    try {
      return switch (target) {
        case SheetSelector.All _ ->
            new WorkbookInspectionResult.SheetsResult(stepId, workbook.sheets().sheetNames());
        case SheetSelector.ByName byName ->
            new WorkbookInspectionResult.SheetsResult(
                stepId,
                workbook.sheets().sheetNames().contains(byName.name())
                    ? List.of(byName.name())
                    : List.of());
        case SheetSelector.ByNames byNames ->
            new WorkbookInspectionResult.SheetsResult(
                stepId,
                byNames.names().stream()
                    .filter(name -> workbook.sheets().sheetNames().contains(name))
                    .toList());
        case NamedRangeSelector selector ->
            executeObservation(
                stepId,
                selector,
                new WorkbookIntrospectionQuery.GetNamedRanges(),
                workbook,
                workbookLocation);
        case TableSelector selector ->
            executeObservation(
                stepId,
                selector,
                new WorkbookAssetIntrospectionQuery.GetTables(),
                workbook,
                workbookLocation);
        case PivotTableSelector selector ->
            executeObservation(
                stepId,
                selector,
                new WorkbookAssetIntrospectionQuery.GetPivotTables(),
                workbook,
                workbookLocation);
        case ChartSelector selector ->
            executeObservation(
                stepId,
                selector,
                new WorkbookAssetIntrospectionQuery.GetCharts(),
                workbook,
                workbookLocation);
        default ->
            throw new IllegalArgumentException(
                "Unsupported presence assertion target: "
                    + SelectorJsonSupport.typeIdsFor(target.getClass()).getFirst());
      };
    } catch (NamedRangeNotFoundException | SheetNotFoundException ignored) {
      return zeroMatchPresenceObservation(stepId, target);
    }
  }

  static InspectionResult zeroMatchPresenceObservation(String stepId, Selector target) {
    return switch (target) {
      case NamedRangeSelector _ ->
          new WorkbookInspectionResult.NamedRangesResult(stepId, List.of());
      case TableSelector _ -> new WorkbookAssetInspectionResult.TablesResult(stepId, List.of());
      case PivotTableSelector _ ->
          new WorkbookAssetInspectionResult.PivotTablesResult(stepId, List.of());
      case ChartSelector.AllOnSheet selector ->
          new WorkbookAssetInspectionResult.ChartsResult(stepId, selector.sheetName(), List.of());
      case ChartSelector.ByName selector ->
          new WorkbookAssetInspectionResult.ChartsResult(stepId, selector.sheetName(), List.of());
      default ->
          throw new IllegalArgumentException(
              "Unsupported presence assertion target: "
                  + SelectorJsonSupport.displayName(target.getClass()));
    };
  }

  static int observedCount(InspectionResult observation) {
    return switch (observation) {
      case WorkbookInspectionResult.SheetsResult result -> result.sheetNames().size();
      case WorkbookInspectionResult.NamedRangesResult result -> result.namedRanges().size();
      case WorkbookAssetInspectionResult.TablesResult result -> result.tables().size();
      case WorkbookAssetInspectionResult.PivotTablesResult result -> result.pivotTables().size();
      case WorkbookAssetInspectionResult.ChartsResult result -> result.charts().size();
      default ->
          throw new IllegalArgumentException(
              "Unsupported presence observation result: " + observation.getClass().getSimpleName());
    };
  }
}
