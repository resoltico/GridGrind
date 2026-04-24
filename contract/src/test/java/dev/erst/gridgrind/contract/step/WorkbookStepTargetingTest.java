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
        List.of(
            new SelectorJsonSupport.FamilyInfo("TableSelector", List.of("TABLE_BY_NAME_ON_SHEET"))),
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
  void exposesFamiliesForDirectSelectorAssertionsWithoutDisambiguationNotes() {
    WorkbookStepTargeting.TargetSurface targetSurface =
        WorkbookStepTargeting.forAssertionType(Assertion.TablePresent.class);

    assertEquals(
        List.of(
            new SelectorJsonSupport.FamilyInfo(
                "TableSelector",
                List.of("TABLE_ALL", "TABLE_BY_NAME", "TABLE_BY_NAMES", "TABLE_BY_NAME_ON_SHEET"))),
        targetSurface.selectorFamilies());
    assertEquals(null, targetSurface.rule());
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
