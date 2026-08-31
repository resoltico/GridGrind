package dev.erst.gridgrind.architecture;

import com.tngtech.archunit.ArchConfiguration;
import com.tngtech.archunit.core.importer.ImportOption;
import com.tngtech.archunit.junit.AnalyzeClasses;
import com.tngtech.archunit.junit.ArchTest;
import com.tngtech.archunit.junit.ArchTests;

/** Executes the complete bytecode-level GridGrind product architecture contract. */
@SuppressWarnings({"PMD.TestClassWithoutTestCases", "PMD.UseUtilityClass"})
@AnalyzeClasses(
    locations = GridGrindProductLocations.class,
    importOptions = {
      ImportOption.DoNotIncludeTests.class,
      ImportOption.DoNotIncludeGradleTestFixtures.class
    })
final class ProductArchitectureTest {
  static {
    if (!"true".equals(ArchConfiguration.get().getProperty("archRule.failOnEmptyShould"))) {
      throw new IllegalStateException(
          "ArchUnit must fail rules whose selected population is empty");
    }
  }

  @ArchTest
  static final ArchTests PRODUCT_ARCHITECTURE = ArchTests.in(ProductArchitectureRules.class);
}
