package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Executes validated workbook read commands in order against one workbook instance. */
public final class WorkbookReadExecutor {
  private final ExcelWorkbookIntrospector workbookIntrospector;
  private final WorkbookAnalyzer analyzer;

  /** Creates the production workbook read executor. */
  public WorkbookReadExecutor() {
    this(new ExcelWorkbookIntrospector(), new WorkbookAnalyzer());
  }

  WorkbookReadExecutor(ExcelWorkbookIntrospector workbookIntrospector, WorkbookAnalyzer analyzer) {
    this.workbookIntrospector =
        Objects.requireNonNull(workbookIntrospector, "workbookIntrospector must not be null");
    this.analyzer = Objects.requireNonNull(analyzer, "analyzer must not be null");
  }

  /** Executes one or more read commands in order and returns their immutable results. */
  public List<WorkbookReadResult> apply(ExcelWorkbook workbook, WorkbookReadCommand... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return apply(workbook, new WorkbookLocation.UnsavedWorkbook(), Arrays.asList(commands));
  }

  /** Executes one or more read commands using the workbook filesystem location for analysis. */
  public List<WorkbookReadResult> apply(
      ExcelWorkbook workbook, WorkbookLocation workbookLocation, WorkbookReadCommand... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return apply(workbook, workbookLocation, Arrays.asList(commands));
  }

  /** Executes read commands from any iterable source in order. */
  public List<WorkbookReadResult> apply(
      ExcelWorkbook workbook, Iterable<WorkbookReadCommand> commands) {
    return apply(workbook, new WorkbookLocation.UnsavedWorkbook(), commands);
  }

  /** Executes read commands from any iterable source with explicit workbook-location context. */
  public List<WorkbookReadResult> apply(
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      Iterable<WorkbookReadCommand> commands) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    List<WorkbookReadResult> results = new ArrayList<>();
    for (WorkbookReadCommand command : commands) {
      Objects.requireNonNull(command, "command must not be null");
      results.add(applyOne(workbook, workbookLocation, command));
    }
    return List.copyOf(results);
  }

  private WorkbookReadResult applyOne(
      ExcelWorkbook workbook, WorkbookLocation workbookLocation, WorkbookReadCommand command) {
    return switch (command) {
      case WorkbookReadCommand.Introspection introspection ->
          workbookIntrospector.execute(workbook, introspection);
      case WorkbookReadCommand.Analysis analysis ->
          analyzer.execute(workbook, workbookLocation, analysis);
    };
  }
}
