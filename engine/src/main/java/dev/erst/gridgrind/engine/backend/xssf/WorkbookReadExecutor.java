package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.plan.InspectionResult;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Executes validated workbook read commands in order against one workbook instance. */
final class WorkbookReadExecutor {
  private final WorkbookIntrospector workbookIntrospector;
  private final WorkbookAnalyzer analyzer;

  /** Creates the production workbook read executor. */
  WorkbookReadExecutor() {
    this(new WorkbookIntrospector(), new WorkbookAnalyzer());
  }

  WorkbookReadExecutor(WorkbookIntrospector workbookIntrospector, WorkbookAnalyzer analyzer) {
    this.workbookIntrospector =
        Objects.requireNonNull(workbookIntrospector, "workbookIntrospector must not be null");
    this.analyzer = Objects.requireNonNull(analyzer, "analyzer must not be null");
  }

  /** Executes one or more read commands in order and returns their immutable results. */
  List<InspectionResult> apply(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedInspectionStep... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return apply(workbook, new WorkbookLocation.UnsavedWorkbook(), Arrays.asList(commands));
  }

  /** Executes one or more read commands using the workbook filesystem location for analysis. */
  List<InspectionResult> apply(
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.core.plan.ResolvedInspectionStep... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return apply(workbook, workbookLocation, Arrays.asList(commands));
  }

  /** Executes read commands from any iterable source in order. */
  List<InspectionResult> apply(
      XssfWorkbook workbook,
      Iterable<dev.erst.gridgrind.core.plan.ResolvedInspectionStep> commands) {
    return apply(workbook, new WorkbookLocation.UnsavedWorkbook(), commands);
  }

  /** Executes read commands from any iterable source with explicit workbook-location context. */
  List<InspectionResult> apply(
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      Iterable<dev.erst.gridgrind.core.plan.ResolvedInspectionStep> commands) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    List<InspectionResult> results = new ArrayList<>();
    for (dev.erst.gridgrind.core.plan.ResolvedInspectionStep command : commands) {
      Objects.requireNonNull(command, "command must not be null");
      results.add(applyOne(workbook, workbookLocation, command));
    }
    return List.copyOf(results);
  }

  private InspectionResult applyOne(
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.core.plan.ResolvedInspectionStep command) {
    return switch (command.query()) {
      case dev.erst.gridgrind.core.plan.InspectionQuery.Introspection _ ->
          workbookIntrospector.execute(workbook, command);
      case dev.erst.gridgrind.core.plan.InspectionQuery.Surface _ ->
          workbookIntrospector.execute(workbook, command);
      case dev.erst.gridgrind.core.plan.InspectionQuery.Analysis _ ->
          analyzer.execute(workbook, workbookLocation, command);
    };
  }
}
