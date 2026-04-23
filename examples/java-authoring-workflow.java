import dev.erst.gridgrind.authoring.GridGrindPlan;
import dev.erst.gridgrind.authoring.Tables;
import dev.erst.gridgrind.authoring.Targets;
import dev.erst.gridgrind.authoring.Values;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import java.nio.file.Path;
import java.util.List;

/** Compile-verified example showing the Java-first authoring API over the canonical contract. */
final class JavaAuthoringWorkflowExample {
  private JavaAuthoringWorkflowExample() {}

  /**
   * Builds one selector-first workflow against a caller-owned workspace.
   *
   * <p>The caller should place one UTF-8 file at
   * {@code workspace/authored-inputs/item.txt} containing the row-key value.
   */
  public static GridGrindPlan build(Path workspace) {
    return GridGrindPlan.newWorkbook()
        .saveAs(workspace.resolve("budget.xlsx"))
        .journal(ExecutionJournalLevel.VERBOSE)
        .mutate(Targets.sheet("Budget").ensureExists())
        .mutate(
            Targets.range("Budget", "A1:B3")
                .setRows(
                    List.of(
                        Values.row(Values.text("Item"), Values.text("Amount")),
                        Values.row(Values.text("Hosting"), Values.number(100.0)),
                        Values.row(Values.text("Travel"), Values.number(50.0)))))
        .mutate(
            Targets.tableOnSheet("BudgetTable", "Budget")
                .define(
                    Tables.define(
                        "BudgetTable",
                        "Budget",
                        "A1:B3",
                        false,
                        Tables.noStyle())))
        .mutate(
            Targets.table("BudgetTable")
                .rowByKey(
                    "Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .set(Values.number(125.0)))
        .inspect(
            Targets.table("BudgetTable")
                .rowByKey(
                    "Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .read())
        .assertThat(
            Targets.table("BudgetTable")
                .rowByKey(
                    "Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .valueEquals(Values.expectedNumber(125.0)));
  }

  /** Executes the authored workflow in-process against the canonical executor. */
  public static GridGrindResponse run(Path workspace) throws Exception {
    return new DefaultGridGrindRequestExecutor()
        .execute(
            build(workspace).toPlan(),
            new ExecutionInputBindings(workspace, (byte[]) null),
            ExecutionJournalSink.NOOP);
  }
}
