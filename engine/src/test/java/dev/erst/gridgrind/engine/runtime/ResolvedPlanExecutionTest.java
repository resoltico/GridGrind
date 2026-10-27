package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.PreparedMutationStep;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises the materialized-plan boundary against actual mutation, persistence and secret use. */
class ResolvedPlanExecutionTest {
  @TempDir Path root;

  @Test
  void materializesFileOperandsButResolvesProtectionSecretsOnlyWhenExecuted() throws Exception {
    Path textFile = Files.writeString(root.resolve("title.txt"), "Prepared title");
    var secret = new SecretReference("workbook-password");
    var protection =
        new WorkbookMutationAction.SetWorkbookProtection(
            new WorkbookProtectionSpec(true, false, false, Optional.of(secret), Optional.empty()));
    var request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("title"),
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(new TextSourceSpec.Utf8File("title.txt")))),
                new MutationStep(
                    new StepId("protect"), new WorkbookSelector.Current(), protection)));
    var resolutions = new AtomicInteger();
    char[] supplied = "private-password".toCharArray();
    var baseGrant = ExecutionGrantTestSupport.forPlan(request, root, textFile);
    var grant =
        new dev.erst.gridgrind.core.grant.ExecutionGrant(
            baseGrant.readableResources(),
            baseGrant.operationIds(),
            baseGrant.targetAuthority(),
            baseGrant.publicationAuthority(),
            List.of(secret),
            baseGrant.acceptancePolicy());
    var bindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            grant,
            reference -> {
              assertEquals(secret, reference);
              resolutions.incrementAndGet();
              return supplied;
            });
    var preflight = RequestPreflight.verify(request, bindings);
    try {
      assertEquals(List.of(), preflight.problems());
      assertEquals(0, resolutions.get());
      var plan = preflight.preparedPlan();
      var text = assertInstanceOf(PreparedMutationStep.class, plan.steps().getFirst());
      var cell = assertInstanceOf(ResolvedMutationAction.SetCell.class, text.action());
      assertEquals(new CellValue.Text("Prepared title"), cell.value());
      var pendingProtection = assertInstanceOf(PreparedMutationStep.class, plan.steps().get(1));
      assertSame(protection, pendingProtection.action());
      Files.delete(textFile);
      var executor =
          new FormulaAwareMutationStepExecutor(
              new SemanticSelectorResolver(new WorkbookExecutionEngine()),
              preflight.preparedBindings());
      Path saved = root.resolve("prepared.xlsx");
      try (var workbook = ExcelWorkbooks.create()) {
        workbook.getOrCreateSheet("Budget");
        var origins = new FormulaOriginTracker();
        executor.execute(
            workbook,
            text,
            origins,
            ExecutionStepContextFactory.referenceFor(0, text.authoredStep()));
        assertEquals("Prepared title", workbook.sheet("Budget").cells().text("A1"));
        assertEquals(0, resolutions.get());
        executor.execute(
            workbook,
            pendingProtection,
            origins,
            ExecutionStepContextFactory.referenceFor(1, pendingProtection.authoredStep()));
        assertEquals(1, resolutions.get());
        assertArrayEquals(new char[supplied.length], supplied);
        workbook
            .persistence()
            .savePlainWorkbook(saved, WorkbookArtifactWriteDisposition.CREATE_NEW);
      }
      try (var reopened = new XSSFWorkbook(saved.toFile())) {
        assertTrue(reopened.isStructureLocked());
        assertTrue(reopened.validateWorkbookPassword("private-password"));
        assertEquals(
            "Prepared title",
            reopened.getSheet("Budget").getRow(0).getCell(0).getStringCellValue());
      }
    } finally {
      preflight.release();
    }
  }
}
