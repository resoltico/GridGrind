package dev.erst.gridgrind.architecture;

import com.tngtech.archunit.core.importer.Location;
import com.tngtech.archunit.junit.LocationProvider;
import dev.erst.gridgrind.authoring.GridGrindPlan;
import dev.erst.gridgrind.cli.GridGrindCli;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import java.security.CodeSource;
import java.util.Objects;
import java.util.Set;

/** Supplies exactly the compiled locations that make up GridGrind's production module graph. */
public final class GridGrindProductLocations implements LocationProvider {
  private static final Set<Class<?>> PRODUCT_MODULE_ANCHORS =
      Set.of(
          GridGrindPlan.class,
          GridGrindCli.class,
          WorkbookPlan.class,
          GridGrindEngine.class,
          ExcelComparisonOperator.class);

  @Override
  public Set<Location> get(Class<?> locationProviderClass) {
    return productLocations();
  }

  static Set<Location> productLocations() {
    return PRODUCT_MODULE_ANCHORS.stream()
        .map(GridGrindProductLocations::locationFor)
        .collect(java.util.stream.Collectors.toUnmodifiableSet());
  }

  static Location locationFor(Class<?> moduleAnchor) {
    CodeSource codeSource =
        Objects.requireNonNull(
            moduleAnchor.getProtectionDomain().getCodeSource(),
            "Every product module anchor must have a compiled code source");
    return Location.of(codeSource.getLocation());
  }
}
