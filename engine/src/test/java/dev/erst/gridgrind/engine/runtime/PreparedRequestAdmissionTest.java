package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.poi.NativeInvocationContext;
import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** The sole prepared-capability factory admits only statically valid, authorized requests. */
class PreparedRequestAdmissionTest {
  @TempDir Path root;

  @Test
  void invalidMutationCannotAcquireInputsOrMintAPreparedCapabilityDespiteAnAdequateGrant()
      throws Exception {
    var plan =
        plan(
            List.of(
                new MutationStep(
                    new StepId("invalid-write"),
                    new WorkbookSelector.Current(),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(TextSourceSpec.standardInput())))));
    var opens = new AtomicInteger();
    var closes = new AtomicInteger();
    var resolutions = new AtomicInteger();
    var inputs = inputs(opens, closes);
    try {
      var result =
          RequestPreflight.verify(
              plan,
              inputs,
              (step, bindings) -> {
                resolutions.incrementAndGet();
                return SourceBackedPlanResolver.resolveStep(step, bindings);
              });
      assertTrue(result.preparedRequest().isEmpty());
      assertEquals(
          List.of(GridGrindProblemCode.INVALID_REQUEST),
          result.problems().stream().map(problem -> problem.code()).toList());
      assertThrows(IllegalStateException.class, result::requirePreparedRequest);
      assertEquals(0, resolutions.get(), "Static rejection must precede input resolution");
      assertEquals(0, opens.get(), "Static rejection must not invoke the lazy input capability");
      assertEquals(0, closes.get());
      assertFalse(Files.exists(root.resolve("scratch")));
      var report =
          new GridGrindRequestDoctor(
                  NativeInvocationContext.enterOwned(
                      dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
              .diagnose(plan, inputs);
      assertEquals(
          result.problems(), report.problems(), "Doctor must not duplicate static findings");
      assertEquals(0, opens.get(), "Doctor must not acquire inputs for a statically invalid plan");
    } finally {
      inputs.closeRunResources();
    }
  }

  @Test
  void validAuthorizedPlanResolvesItsPayloadOnceBeforeMintingTheCapability() throws Exception {
    var plan =
        plan(
            List.of(
                new MutationStep(
                    new StepId("sheet"),
                    new SheetSelector.ByName(new SheetName("Data")),
                    new WorkbookMutationAction.EnsureSheet()),
                new MutationStep(
                    new StepId("write"),
                    new CellSelector.ByAddress(new SheetName("Data"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(TextSourceSpec.standardInput())))));
    var opens = new AtomicInteger();
    var closes = new AtomicInteger();
    var resolutions = new AtomicInteger();
    var inputs = inputs(opens, closes);
    try {
      var result =
          RequestPreflight.verify(
              plan,
              inputs,
              (step, bindings) -> {
                resolutions.incrementAndGet();
                return SourceBackedPlanResolver.resolveStep(step, bindings);
              });
      assertTrue(result.problems().isEmpty(), result.problems().toString());
      assertEquals(plan, result.requirePreparedRequest().plan().request());
      assertEquals(2, resolutions.get());
      assertEquals(1, opens.get());
      assertEquals(1, closes.get());
    } finally {
      inputs.closeRunResources();
    }
  }

  private ExecutionInputBindings inputs(AtomicInteger opens, AtomicInteger closes) {
    var grant =
        new ExecutionGrant(
            8,
            List.of(new ExecutionGrant.ReadAuthority.Stdin()),
            List.of("ENSURE_SHEET", "SET_CELL"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());
    return new ExecutionInputBindings(
        root,
        root.resolve("scratch"),
        grant,
        Optional.of(() -> {
          opens.incrementAndGet();
          return new ByteArrayInputStream("accepted".getBytes(StandardCharsets.UTF_8)) {
            @Override
            public void close() {
              closes.incrementAndGet();
            }
          };
        }),
        Optional.empty());
  }

  private static WorkbookPlan plan(List<WorkbookStep> steps) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        steps);
  }
}
