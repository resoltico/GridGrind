package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.ResolvedCellText;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.streaming.StreamingWorkbookWriter;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.poi.NativeInvocationContext;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies actual materialization counts, session-owned native lifetimes and observation order. */
class StreamingReadViewTest {
  @TempDir Path root;

  @Test
  void terminalObservationsShareOneMaterializationAndNativeOpen() throws Exception {
    var engine = new WorkbookExecutionEngine();
    var selectors = new SemanticSelectorResolver(engine);
    var bindings =
        new ExecutionInputBindings(
            root, root.resolve("scratch"), ExecutionGrantTestSupport.noPublication());
    var opens = new AtomicInteger();
    var nativeOpened =
        new java.util.concurrent.atomic.AtomicReference<
            dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook>();
    var snapshot = new java.util.concurrent.atomic.AtomicReference<Path>();
    TempFileFactory temp = (prefix, suffix) -> Files.createTempFile(root, prefix, suffix);
    var steps =
        new StepExecution(
            engine,
            selectors,
            new AssertionExecutor(engine, selectors),
            temp,
            bindings,
            path -> {
              opens.incrementAndGet();
              var opened = XssfWorkbooks.open(path, temp::createTempFile);
              nativeOpened.set(opened);
              snapshot.set(path);
              return opened;
            });
    var sheet = new SheetSelector.ByName(new SheetName("Data"));
    try (var writer = new StreamingWorkbookWriter()) {
      writer.apply(new ResolvedStep(sheet, new WorkbookMutationAction.EnsureSheet()));
      writer.apply(
          new ResolvedStep(
              sheet, new ResolvedMutationAction.AppendRow(List.of(new ResolvedCellText("one")))));
      try (var view = steps.streamingReadView(writer, new WorkbookLocation.UnsavedWorkbook())) {
        assertInstanceOf(
            AssertionResult.Passed.class,
            view.assertThat(
                (dev.erst.gridgrind.core.plan.PreparedAssertionStep)
                    dev.erst.gridgrind.core.plan.PreparedStep.prepare(assertion("A1", "one"))));
        view.inspect(
            (dev.erst.gridgrind.core.plan.PreparedInspectionStep)
                dev.erst.gridgrind.core.plan.PreparedStep.prepare(
                    new InspectionStep(
                        new StepId("inspect"),
                        new CellSelector.ByAddress(new SheetName("Data"), new CellAddress("A1")),
                        new SheetIntrospectionQuery.GetCells())));
        assertEquals(
            1, opens.get(), "Same-state assertions and inspections share one actual XSSF open");
        assertInstanceOf(
            AssertionResult.Passed.class,
            view.assertThat(
                (dev.erst.gridgrind.core.plan.PreparedAssertionStep)
                    dev.erst.gridgrind.core.plan.PreparedStep.prepare(assertion("A1", "one"))));
        assertEquals(1, opens.get(), "Terminal observations share the same native workbook");
        // These assertions execute before the fallback TWR close, so that close cannot mask a
        // session leak.
        assertNotNull(nativeOpened.get().xssfWorkbook().getPackage());
        assertTrue(Files.isRegularFile(snapshot.get()));
        bindings.closeRunResources();
        assertNull(nativeOpened.get().xssfWorkbook().getPackage());
        assertFalse(Files.exists(snapshot.get()));
      }
    } finally {
      bindings.closeRunResources();
    }
    try (var files = Files.list(root)) {
      assertEquals(
          0,
          files
              .filter(path -> path.getFileName().toString().startsWith("gridgrind-streaming-step-"))
              .count());
    }
  }

  @Test
  void interleavedStreamingRequestUsesFullBackendAndPreservesEachObservedState() throws Exception {
    var sheet = new SheetSelector.ByName(new SheetName("Data"));
    var plan =
        dev.erst.gridgrind.core.plan.WorkbookPlan.standard(
            new dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookSource.New(),
            new dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicySpec.mode(
                new dev.erst.gridgrind.core.plan.ExecutionModeSpec.StreamingWrite()),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
            List.of(
                new dev.erst.gridgrind.core.plan.MutationStep(
                    new StepId("sheet"), sheet, new WorkbookMutationAction.EnsureSheet()),
                new dev.erst.gridgrind.core.plan.MutationStep(
                    new StepId("row-one"),
                    sheet,
                    new dev.erst.gridgrind.core.plan.CellMutationAction.AppendRow(
                        new dev.erst.gridgrind.core.model.value.CellRowSpec.Typed(
                            List.of(
                                new dev.erst.gridgrind.core.model.value.CellValueSpec.Text(
                                    dev.erst.gridgrind.core.source.TextSourceSpec.inline(
                                        "one")))))),
                new InspectionStep(
                    new StepId("before-second-row"),
                    new CellSelector.ByAddresses(
                        new SheetName("Data"),
                        List.of(new CellAddress("A1"), new CellAddress("A2"))),
                    new SheetIntrospectionQuery.GetCells()),
                new dev.erst.gridgrind.core.plan.MutationStep(
                    new StepId("row-two"),
                    sheet,
                    new dev.erst.gridgrind.core.plan.CellMutationAction.AppendRow(
                        new dev.erst.gridgrind.core.model.value.CellRowSpec.Typed(
                            List.of(
                                new dev.erst.gridgrind.core.model.value.CellValueSpec.Text(
                                    dev.erst.gridgrind.core.source.TextSourceSpec.inline(
                                        "two")))))),
                new InspectionStep(
                    new StepId("after-second-row"),
                    new CellSelector.ByAddresses(
                        new SheetName("Data"),
                        List.of(new CellAddress("A1"), new CellAddress("A2"))),
                    new SheetIntrospectionQuery.GetCells())));
    var response =
        new DefaultGridGrindRequestExecutor(
                NativeInvocationContext.enterOwned(
                    dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
            .execute(plan, ExecutionInputBindingsFixtureSupport.bindings(root, plan));
    var success =
        assertInstanceOf(
            dev.erst.gridgrind.core.result.WorkbookResult.Success.class,
            response,
            response.toString());
    assertEquals(
        dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Backend.FULL_XSSF,
        success.evidence().admission().executionStrategy().orElseThrow().backend());
    var before =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult.class,
            success.inspections().get(0));
    var after =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult.class,
            success.inspections().get(1));
    assertEquals(2, before.cells().size());
    assertEquals(
        "one",
        assertInstanceOf(
                dev.erst.gridgrind.core.model.value.CellReport.TextReport.class,
                before.cells().get(0))
            .value()
            .orElseThrow()
            .value());
    assertInstanceOf(
        dev.erst.gridgrind.core.model.value.CellReport.BlankReport.class, before.cells().get(1));
    assertEquals(
        List.of("one", "two"),
        after.cells().stream()
            .map(cell ->
                assertInstanceOf(
                        dev.erst.gridgrind.core.model.value.CellReport.TextReport.class, cell)
                    .value()
                    .orElseThrow()
                    .value())
            .toList());
  }

  private static AssertionStep assertion(String address, String text) {
    return new AssertionStep(
        new StepId("check-" + address),
        new CellSelector.ByAddress(new SheetName("Data"), new CellAddress(address)),
        new CellAssertion.Value(new CellValue.Text(text)));
  }
}
