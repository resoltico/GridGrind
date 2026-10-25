package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.catalog.OperationEffect;
import dev.erst.gridgrind.contract.catalog.OperationEffectFootprint;
import dev.erst.gridgrind.contract.selector.CellSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies the typed execution-evidence states reject ambiguous or malformed claims. */
class WorkbookExecutionEvidenceTest {
  @Test
  void minimumKeepsEvidenceClaimsDistinct() {
    WorkbookExecutionEvidence evidence = WorkbookExecutionEvidence.minimum();

    assertInstanceOf(WorkbookExecutionEvidence.Admission.class, evidence.admission());
    assertInstanceOf(
        WorkbookExecutionEvidence.Structural.NotEstablished.class, evidence.structural());
    assertInstanceOf(
        WorkbookExecutionEvidence.Computational.NotRequested.class, evidence.computational());
    assertInstanceOf(
        WorkbookExecutionEvidence.TaskSpecific.NotRequested.class, evidence.taskSpecific());
    assertInstanceOf(
        WorkbookExecutionEvidence.Presentational.NotAssessed.class, evidence.presentational());
    assertInstanceOf(
        WorkbookExecutionEvidence.Preservation.NotRequested.class,
        evidence.preservation().getFirst());
  }

  @Test
  void evidenceStatesRetainDistinctEstablishedFailedAndPreservationFacts() {
    WorkbookExecutionEvidence.Operation operation =
        new WorkbookExecutionEvidence.Operation(
            "SET_CELL", List.of(OperationEffect.CELLS), OperationEffectFootprint.TARGET_BOUNDED);
    WorkbookExecutionEvidence.MaterializedInput input =
        new WorkbookExecutionEvidence.MaterializedInput("SOURCE", "input.xlsx", 3, "0".repeat(64));
    AssertionResult.Passed passed = new AssertionResult.Passed("assert-cell", "EXPECT_CELL_VALUE");
    AssertionResult.Failed failed =
        new AssertionResult.Failed(
            "assert-cell",
            "EXPECT_CELL_VALUE",
            new AssertionFailure(
                "assert-cell",
                "EXPECT_CELL_VALUE",
                new CellSelector.ByAddress("Summary", "A1"),
                new CellAssertion.CellValue(new CellScalarValue.Text("Approved")),
                List.of()));
    WorkbookExecutionEvidence.TaskSpecific.Assertion passedAssertion =
        new WorkbookExecutionEvidence.TaskSpecific.Assertion(
            WorkbookExecutionEvidence.TaskSpecific.Origin.PLAN_AUTHORED, passed);
    WorkbookExecutionEvidence.TaskSpecific.Assertion failedAssertion =
        new WorkbookExecutionEvidence.TaskSpecific.Assertion(
            WorkbookExecutionEvidence.TaskSpecific.Origin.HOST_REQUIRED, failed);
    WorkbookExecutionEvidence evidence =
        new WorkbookExecutionEvidence(
            new WorkbookExecutionEvidence.Admission(List.of(operation), List.of(input)),
            new WorkbookExecutionEvidence.Structural.Established(),
            new WorkbookExecutionEvidence.Computational.Established(),
            new WorkbookExecutionEvidence.TaskSpecific.Established(List.of(passedAssertion)),
            new WorkbookExecutionEvidence.Presentational.NotAssessed(),
            List.of(
                new WorkbookExecutionEvidence.Preservation.SemanticEstablished(
                    List.of("preserve-summary")),
                new WorkbookExecutionEvidence.Preservation.ByteEstablished(
                    List.of("/customXml/item1.xml")),
                new WorkbookExecutionEvidence.Preservation.Failed(
                    WorkbookExecutionEvidence.Preservation.Kind.BYTE,
                    List.of("/customXml/item2.xml")),
                new WorkbookExecutionEvidence.Preservation.Unestablished()));

    assertEquals(List.of(operation), evidence.admission().operations());
    assertInstanceOf(
        WorkbookExecutionEvidence.TaskSpecific.Established.class, evidence.taskSpecific());
    assertInstanceOf(
        WorkbookExecutionEvidence.Preservation.SemanticEstablished.class,
        evidence.preservation().getFirst());
    assertInstanceOf(
        WorkbookExecutionEvidence.TaskSpecific.Failed.class,
        new WorkbookExecutionEvidence.TaskSpecific.Failed(List.of(failedAssertion)));
  }

  @Test
  void evidenceRejectsEmptyInvalidOrSelfContradictoryClaims() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookExecutionEvidence.Operation(
                "SET_CELL", List.of(), OperationEffectFootprint.TARGET_BOUNDED));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookExecutionEvidence.MaterializedInput(
                "SOURCE", "input.xlsx", -1, "0".repeat(64)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookExecutionEvidence.TaskSpecific.Established(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookExecutionEvidence.TaskSpecific.Failed(
                List.of(
                    new WorkbookExecutionEvidence.TaskSpecific.Assertion(
                        WorkbookExecutionEvidence.TaskSpecific.Origin.PLAN_AUTHORED,
                        new AssertionResult.Passed("assert-cell", "EXPECT_CELL_VALUE")))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookExecutionEvidence.TaskSpecific.Established(List.of(failedAssertion())));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookExecutionEvidence.Preservation.SemanticEstablished(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookExecutionEvidence.Preservation.ByteEstablished(List.of(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookExecutionEvidence.MaterializedInput("SOURCE", "input.xlsx", 1, "invalid"));
    assertInstanceOf(
        WorkbookExecutionEvidence.Preservation.NotApplicable.class,
        new WorkbookExecutionEvidence.Preservation.NotApplicable());
  }

  private static WorkbookExecutionEvidence.TaskSpecific.Assertion failedAssertion() {
    AssertionFailure failure =
        new AssertionFailure(
            "assert-cell",
            "EXPECT_CELL_VALUE",
            new CellSelector.ByAddress("Summary", "A1"),
            new CellAssertion.CellValue(new CellScalarValue.Text("Approved")),
            List.of());
    return new WorkbookExecutionEvidence.TaskSpecific.Assertion(
        WorkbookExecutionEvidence.TaskSpecific.Origin.HOST_REQUIRED,
        new AssertionResult.Failed("assert-cell", "EXPECT_CELL_VALUE", failure));
  }
}
