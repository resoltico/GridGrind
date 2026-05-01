package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.WorkbookReadResult;
import java.util.List;
import java.util.Objects;

/** Shared assertion observation adapter over selector resolution and workbook read execution. */
final class AssertionObservationExecutor {
  private final WorkbookReadExecutor readExecutor;
  private final SemanticSelectorResolver selectorResolver;

  AssertionObservationExecutor(
      WorkbookReadExecutor readExecutor, SemanticSelectorResolver selectorResolver) {
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
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
      return resolvedTarget.shortCircuitResult();
    }
    WorkbookReadResult result =
        readExecutor
            .apply(
                workbook,
                workbookLocation,
                InspectionCommandConverter.toReadCommand(stepId, resolvedTarget.selector(), query))
            .getFirst();
    return InspectionResultConverter.toReadResult(result);
  }

  InspectionResult presenceObservation(
      String stepId, Selector target, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    try {
      return switch (target) {
        case NamedRangeSelector selector ->
            executeObservation(
                stepId, selector, new InspectionQuery.GetNamedRanges(), workbook, workbookLocation);
        case TableSelector selector ->
            executeObservation(
                stepId, selector, new InspectionQuery.GetTables(), workbook, workbookLocation);
        case PivotTableSelector selector ->
            executeObservation(
                stepId, selector, new InspectionQuery.GetPivotTables(), workbook, workbookLocation);
        case ChartSelector _ -> chartsObservation(stepId, target, workbook, workbookLocation);
        default ->
            throw new IllegalArgumentException(
                "Unsupported presence assertion target: " + target.getClass().getSimpleName());
      };
    } catch (NamedRangeNotFoundException | SheetNotFoundException ignored) {
      return zeroMatchPresenceObservation(stepId, target);
    }
  }

  InspectionResult.ChartsResult chartsObservation(
      String stepId, Selector target, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    String sheetName =
        switch (target) {
          case ChartSelector.AllOnSheet selector -> selector.sheetName();
          case ChartSelector.ByName selector -> selector.sheetName();
          default -> throw new IllegalArgumentException("Unsupported chart assertion target");
        };
    InspectionResult.ChartsResult allCharts =
        (InspectionResult.ChartsResult)
            executeObservation(
                stepId,
                new SheetSelector.ByName(sheetName),
                new InspectionQuery.GetCharts(),
                workbook,
                workbookLocation);
    if (target instanceof ChartSelector.ByName selector) {
      return new InspectionResult.ChartsResult(
          stepId,
          allCharts.sheetName(),
          allCharts.charts().stream()
              .filter(chart -> chart.name().equals(selector.chartName()))
              .toList());
    }
    return allCharts;
  }

  static InspectionResult zeroMatchPresenceObservation(String stepId, Selector target) {
    return switch (target) {
      case NamedRangeSelector _ -> new InspectionResult.NamedRangesResult(stepId, List.of());
      case TableSelector _ -> new InspectionResult.TablesResult(stepId, List.of());
      case PivotTableSelector _ -> new InspectionResult.PivotTablesResult(stepId, List.of());
      case ChartSelector.AllOnSheet selector ->
          new InspectionResult.ChartsResult(stepId, selector.sheetName(), List.of());
      case ChartSelector.ByName selector ->
          new InspectionResult.ChartsResult(stepId, selector.sheetName(), List.of());
      default ->
          throw new IllegalArgumentException(
              "Unsupported presence assertion target: " + target.getClass().getSimpleName());
    };
  }

  static int observedCount(InspectionResult observation) {
    return switch (observation) {
      case InspectionResult.NamedRangesResult result -> result.namedRanges().size();
      case InspectionResult.TablesResult result -> result.tables().size();
      case InspectionResult.PivotTablesResult result -> result.pivotTables().size();
      case InspectionResult.ChartsResult result -> result.charts().size();
      default ->
          throw new IllegalArgumentException(
              "Unsupported presence observation result: " + observation.getClass().getSimpleName());
    };
  }
}
