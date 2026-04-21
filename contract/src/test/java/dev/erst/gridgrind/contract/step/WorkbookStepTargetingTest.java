package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct tests for public target-selector surface metadata. */
class WorkbookStepTargetingTest {
  @Test
  void exposesStaticSelectorFamiliesForMutations() {
    WorkbookStepTargeting.TargetSurface targetSurface =
        WorkbookStepTargeting.forMutationActionType(MutationAction.SetTable.class);

    assertEquals(
        List.of(new SelectorJsonSupport.FamilyInfo("TableSelector", List.of("BY_NAME_ON_SHEET"))),
        targetSurface.selectorFamilies());
    assertEquals(null, targetSurface.rule());
  }

  @Test
  void exposesDerivedRulesForDynamicAssertions() {
    WorkbookStepTargeting.TargetSurface targetSurface =
        WorkbookStepTargeting.forAssertionType(Assertion.AnalysisFindingPresent.class);

    assertTrue(targetSurface.selectorFamilies().isEmpty());
    assertEquals("Matches the nested analysis query's target selectors.", targetSurface.rule());
  }

  @Test
  void exposesFamiliesAndDisambiguationRuleForSharedSelectorAssertions() {
    WorkbookStepTargeting.TargetSurface targetSurface =
        WorkbookStepTargeting.forAssertionType(Assertion.Present.class);

    assertEquals(
        List.of(
            new SelectorJsonSupport.FamilyInfo(
                "NamedRangeSelector",
                List.of("ALL", "ANY_OF", "BY_NAME", "BY_NAMES", "WORKBOOK_SCOPE", "SHEET_SCOPE")),
            new SelectorJsonSupport.FamilyInfo(
                "TableSelector", List.of("ALL", "BY_NAME", "BY_NAMES", "BY_NAME_ON_SHEET")),
            new SelectorJsonSupport.FamilyInfo(
                "PivotTableSelector", List.of("ALL", "BY_NAME", "BY_NAMES", "BY_NAME_ON_SHEET")),
            new SelectorJsonSupport.FamilyInfo(
                "ChartSelector", List.of("ALL_ON_SHEET", "BY_NAME"))),
        targetSurface.selectorFamilies());
    assertTrue(targetSurface.rule().contains("Shared selector wire types remain family-sensitive"));
    assertTrue(targetSurface.rule().contains("BY_NAMES"));
    assertTrue(targetSurface.rule().contains("BY_NAME_ON_SHEET"));
  }

  @Test
  void rejectsTargetSurfacesThatDeclareNeitherFamiliesNorRule() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookStepTargeting.TargetSurface(List.of(), null));

    assertEquals(
        "target surface must declare selectorFamilies or a non-blank rule", failure.getMessage());
  }

  @Test
  void rejectsTargetSurfacesWithNullFamiliesOrBlankRule() {
    assertEquals(
        "selectorFamilies must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new WorkbookStepTargeting.TargetSurface(null, "derived"))
            .getMessage());

    assertEquals(
        "rule must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookStepTargeting.TargetSurface(List.of(), " "))
            .getMessage());
  }
}
