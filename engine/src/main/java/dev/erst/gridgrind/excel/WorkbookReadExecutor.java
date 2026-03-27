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
    return apply(workbook, Arrays.asList(commands));
  }

  /** Executes read commands from any iterable source in order. */
  public List<WorkbookReadResult> apply(
      ExcelWorkbook workbook, Iterable<WorkbookReadCommand> commands) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(commands, "commands must not be null");

    List<WorkbookReadResult> results = new ArrayList<>();
    for (WorkbookReadCommand command : commands) {
      Objects.requireNonNull(command, "command must not be null");
      results.add(applyOne(workbook, command));
    }
    return List.copyOf(results);
  }

  private WorkbookReadResult applyOne(ExcelWorkbook workbook, WorkbookReadCommand command) {
    return switch (command) {
      case WorkbookReadCommand.Introspection introspection ->
          workbookIntrospector.execute(workbook, introspection);
      case WorkbookReadCommand.Analysis analysis -> analyzer.execute(workbook, analysis);
    };
  }
}
