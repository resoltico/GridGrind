package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies host acceptance requirements are non-empty, distinct, and exact. */
class GridGrindHostAcceptancePolicyTest {
  @Test
  @SuppressWarnings(
      "PMD.AvoidInstantiatingObjectsInLoops") // Each invalid part name exercises one normalization
  // branch.
  void requirementsRejectAmbiguityAndPreserveValidHostAssertions() {
    AssertionStep assertion =
        new AssertionStep(
            "assert-owner",
            new CellSelector.ByAddress("Budget", "A1"),
            new CellAssertion.CellValue(new CellScalarValue.Text("Ada")));
    InspectionStep inspection =
        new InspectionStep(
            "inspect-budget",
            new SheetSelector.ByName("Budget"),
            new SheetIntrospectionQuery.GetSheetSummary());

    assertEquals(
        List.of(assertion),
        new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(List.of(assertion))
            .assertions());
    assertEquals(
        List.of(inspection),
        new GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts(List.of(inspection))
            .inspections());
    assertEquals(
        List.of("/customXml/item1.xml"),
        new GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                List.of("/customXml/item1.xml"))
            .partNames());

    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindHostAcceptancePolicy.RequireAll(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(
                List.of(assertion, assertion)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts(
                List.of(inspection, inspection)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                List.of("/customXml/item1.xml", "/customXml/item1.xml")));
    for (String invalidPartName :
        List.of(
            "custom.xml",
            "/",
            "/part/",
            "/a\\b",
            "/a//b",
            "/a\u0000b",
            "/.",
            "/..",
            "/a/./b",
            "/a/../b")) {
      assertThrows(
          IllegalArgumentException.class,
          () ->
              new GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                  List.of(invalidPartName)));
    }
  }
}
