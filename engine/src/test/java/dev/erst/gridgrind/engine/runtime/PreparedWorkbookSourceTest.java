package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactIo;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests native preparation handoff and explicit close obligations without masking the owner. */
class PreparedWorkbookSourceTest {
  @TempDir Path root;

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // The factory registers the source with the session closed in finally.
  void eventPreparationUsesSecurityAndMetadataWithoutOpeningAnXssfGraph() throws Exception {
    WorkbookPlan request = request(new ExecutionModeSpec.EventRead());
    var bindings = ExecutionInputBindingsFixtureSupport.bindings(root, request);
    try (var access =
        new RequestPathAccess(root, bindings.tempFileFactory(), bindings.executionGrant())) {
      var preparedBindings = bindings.withRequestPathAccess(access);
      var problems = new ArrayList<GridGrindProblemDetail.Problem>();
      var source =
          RequestPreflightWorkbookSource.prepare(
              request,
              preparedBindings,
              problems,
              (plan, inputs) -> {
                throw new AssertionError("Event-only source opened full XSSF");
              });
      assertEquals(List.of(), problems);
      assertEquals(WorkbookExecutionEvidence.Backend.EVENT_READ, source.strategy().backend());
      assertFalse(source.strategy().fullWorkbookMaterializationRequired());
      try (var materialized = source.takeMaterialized().orElseThrow()) {
        assertTrue(Files.isRegularFile(materialized.workbookPath()));
      }
      assertTrue(source.takeMaterialized().isEmpty());
    } finally {
      bindings.closeRunResources();
    }
  }

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // The factory registers the source with the session closed in finally.
  void fullPreparationTransfersTheSameNativeResourceExactlyOnce() throws Exception {
    WorkbookPlan request = request(new ExecutionModeSpec.FullXssf());
    var bindings = ExecutionInputBindingsFixtureSupport.bindings(root, request);
    var opens = new AtomicInteger();
    var opened = new AtomicReference<XssfWorkbook>();
    try (var access =
        new RequestPathAccess(root, bindings.tempFileFactory(), bindings.executionGrant())) {
      var preparedBindings = bindings.withRequestPathAccess(access);
      var problems = new ArrayList<GridGrindProblemDetail.Problem>();
      var source =
          RequestPreflightWorkbookSource.prepare(
              request,
              preparedBindings,
              problems,
              (plan, inputs) -> {
                opens.incrementAndGet();
                var workbook =
                    new WorkbookLifecycle(inputs.tempFileFactory())
                        .openWorkbook(plan.source(), plan.formulaEnvironment(), inputs);
                opened.set(workbook);
                return workbook;
              });
      assertEquals(List.of(), problems);
      assertEquals(1, opens.get());
      try (var transferred = source.takeXssf().orElseThrow()) {
        assertSame(opened.get(), transferred);
        source.close();
        bindings.closeRunResources();
        // The caller's TWR close has not run: the session must not close the transferred handle.
        assertNotNull(transferred.xssfWorkbook().getPackage());
        assertEquals(1, transferred.sheets().sheetCount());
      }
      assertNull(opened.get().xssfWorkbook().getPackage());
      assertTrue(source.takeXssf().isEmpty());
    } finally {
      bindings.closeRunResources();
    }
  }

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // This case deliberately invokes close twice to test the holder's
  // idempotence.
  void untransferredMaterializationClosesExactlyOnce() throws Exception {
    var closed = new AtomicInteger();
    var strategy =
        new WorkbookExecutionEvidence.ExecutionStrategy(
            WorkbookExecutionEvidence.Backend.EVENT_READ, false);
    var source = new PreparedWorkbookSource(strategy);
    source.retain(
        new WorkbookArtifactIo.MaterializedWorkbook() {
          @Override
          public Path workbookPath() {
            return root;
          }

          @Override
          public void close() {
            closed.incrementAndGet();
          }
        });
    assertEquals(0, closed.get());
    source.close();
    assertEquals(1, closed.get());
    source.close();
    assertEquals(1, closed.get());
  }

  private WorkbookPlan request(ExecutionModeSpec mode) throws java.io.IOException {
    Path source = root.resolve("source.xlsx");
    try (var workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet("Data").createRow(0).createCell(0).setCellFormula("1+1");
      workbook.write(output);
    }
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx"),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicySpec.mode(mode),
        FormulaEnvironmentSpec.empty(),
        List.of());
  }
}
