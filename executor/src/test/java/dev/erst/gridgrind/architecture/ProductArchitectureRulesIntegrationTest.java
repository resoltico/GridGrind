package dev.erst.gridgrind.architecture;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import com.tngtech.archunit.core.importer.ClassFileImporter;
import org.junit.jupiter.api.Test;

/** Exercises every mandatory rule over precisely the compiled production-module locations. */
class ProductArchitectureRulesIntegrationTest {
  @Test
  void everyMandatoryRuleAcceptsTheCurrentCompiledProduct() {
    var productClasses =
        new ClassFileImporter().importLocations(GridGrindProductLocations.productLocations());

    ProductArchitectureRules.mandatoryRules()
        .forEach(
            rule -> assertDoesNotThrow(() -> rule.check(productClasses), rule.getDescription()));
  }
}
