import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.core.authoring.GridGrindPlan;
import dev.erst.gridgrind.core.authoring.ExpectedValues;
import dev.erst.gridgrind.core.authoring.Tables;
import dev.erst.gridgrind.core.authoring.Targets;
import dev.erst.gridgrind.core.authoring.Values;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindProgressSink;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Stream;

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
        .saveAs(workspace.resolve("budget.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REPLACE)
        .journal(ExecutionJournalLevel.VERBOSE)
        .mutate(Targets.sheet(new SheetName("Budget")).ensureExists())
        .mutate(
            Targets.range(new SheetName("Budget"), new RangeRef("A1:B3"))
                .setRows(
                    List.of(
                        Values.row(Values.text("Item"), Values.text("Amount")),
                        Values.row(Values.text("Hosting"), Values.number(100.0)),
                        Values.row(Values.text("Travel"), Values.number(50.0)))))
        .mutate(
            Targets.tableOnSheet(new TableName("BudgetTable"), new SheetName("Budget"))
                .define(
                    Tables.define(
                        new TableName("BudgetTable"),
                        new SheetName("Budget"),
                        new RangeRef("A1:B3"),
                        false,
                        Tables.noStyle())))
        .mutate(
            Targets.table(new TableName("BudgetTable"))
                .rowByKey(
                    "Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .set(Values.number(125.0)))
        .inspect(
            Targets.table(new TableName("BudgetTable"))
                .rowByKey(
                    "Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .read())
        .assertThat(
            Targets.table(new TableName("BudgetTable"))
                .rowByKey(
                    "Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .valueEquals(ExpectedValues.number(125.0)));
  }

  /** Executes the authored workflow in-process against the canonical executor. */
  public static WorkbookResult run(Path workspace) throws Exception {
    Path tempRoot = Files.createTempDirectory("gridgrind-java-");
    try {
      return GridGrindEngine.requestExecutor()
          .execute(
              build(workspace).toPlan(),
              new GridGrindRequestInputs(workspace, tempRoot, executionGrant(workspace)),
              GridGrindProgressSink.NOOP);
    } finally {
      deleteTreeIfExists(tempRoot);
    }
  }

  private static ExecutionGrant executionGrant(Path workspace) {
    return new ExecutionGrant(
        List.of(
            new ExecutionGrant.ReadAuthority.File(
                workspace.resolve("authored-inputs").resolve("item.txt"))),
        List.of(
            "ENSURE_SHEET",
            "SET_RANGE",
            "SET_TABLE",
            "SET_CELL",
            "GET_CELLS",
            "EXPECT_CELL_VALUE"),
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new ExecutionGrant.PublicationAuthority.SaveAs(
            workspace.resolve("budget.xlsx"),
            WorkbookPlan.WorkbookPersistence.IfExists.REPLACE),
        List.of(),
        dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum());
  }

  private static void deleteTreeIfExists(Path root) throws IOException {
    if (!Files.exists(root)) {
      return;
    }
    try (Stream<Path> paths = Files.walk(root)) {
      for (Path path : paths.sorted(Comparator.reverseOrder()).toList()) {
        Files.deleteIfExists(path);
      }
    }
  }
}
