package dev.erst.gridgrind.core.problem;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Checks the parser-independent intake value at the engine handoff boundary. */
class RequestIntakeTest {
  @Test
  @SuppressWarnings("PMD.UseConcurrentHashMap") // Local mutable input proves defensive ownership.
  void ownsTokenOffsetsAndRejectsACompletePlanAlongsideProblems() {
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    RequestBoundFragments fragments =
        new RequestBoundFragments(
            new RequestBoundRoot(
                Optional.of(ProtocolVersion.CURRENT),
                Optional.empty(),
                Optional.of(plan.source()),
                Optional.of(plan.persistence()),
                Optional.of(plan.execution()),
                Optional.of(plan.formulaEnvironment())),
            Optional.of(List.of()));
    Map<String, Long> offsets = new HashMap<>();
    offsets.put("source", 10L);
    RequestIntake intake =
        new RequestIntake(fragments, List.of(), List.of(), Optional.of(plan), offsets, List.of());
    offsets.clear();
    assertEquals(Optional.of(10L), intake.byteOffsetAt("source"));
    assertTrue(intake.byteOffsetAt("missing").isEmpty());
    assertThrows(UnsupportedOperationException.class, () -> intake.tokenOffsets().clear());
    assertEquals(plan, intake.requireCompletePlan());
    assertTrue(intake.isBindable());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestIntake(
                fragments,
                List.of(new RequestUnknownField("extra", 20)),
                List.of(),
                Optional.of(plan),
                Map.of(),
                List.of()));
  }
}
