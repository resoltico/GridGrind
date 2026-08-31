package dev.erst.gridgrind.architecture;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import com.tngtech.archunit.core.importer.Location;
import dev.erst.gridgrind.architecture.fixture.ArchitectureRuntimeLeakFixture;
import dev.erst.gridgrind.authoring.GridGrindPlan;
import dev.erst.gridgrind.cli.GridGrindCli;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Verifies that architecture imports are restricted to the canonical product modules. */
class GridGrindProductLocationsTest {
  @Test
  void importsExactlyTheFiveCanonicalProductModuleLocations() {
    Set<Location> expectedLocations =
        Set.of(
            GridGrindProductLocations.locationFor(GridGrindPlan.class),
            GridGrindProductLocations.locationFor(GridGrindCli.class),
            GridGrindProductLocations.locationFor(WorkbookPlan.class),
            GridGrindProductLocations.locationFor(GridGrindEngine.class),
            GridGrindProductLocations.locationFor(ExcelComparisonOperator.class));

    assertEquals(expectedLocations, GridGrindProductLocations.productLocations());
    assertEquals(expectedLocations, new GridGrindProductLocations().get(getClass()));
  }

  @Test
  void excludesArchitectureRegressionFixturesFromTheProductionImport() {
    Location fixtureLocation =
        GridGrindProductLocations.locationFor(ArchitectureRuntimeLeakFixture.class);

    assertFalse(GridGrindProductLocations.productLocations().contains(fixtureLocation));
  }
}
