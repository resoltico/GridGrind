package dev.erst.gridgrind.architecture;

import static org.junit.jupiter.api.Assertions.assertThrows;

import com.tngtech.archunit.core.importer.ClassFileImporter;
import dev.erst.gridgrind.architecture.fixture.ArchitectureUnadmittedExecutorFixture;
import dev.erst.gridgrind.engine.runtime.ArchitecturePublicationBypassFixture;
import org.junit.jupiter.api.Test;

/** Proves the assurance-boundary rules reject deliberate runtime bypass fixtures. */
class ProductAssuranceArchitectureRulesTest {
  @Test
  void rejectsAnExecutorThatDoesNotCrossTheAdmittedProductionEntryPoint() {
    assertThrows(
        AssertionError.class,
        () ->
            ProductAssuranceArchitectureRules.executionUsesOneAdmittedExecutor()
                .check(
                    new ClassFileImporter()
                        .importClasses(ArchitectureUnadmittedExecutorFixture.class)));
  }

  @Test
  void rejectsASecondRuntimeConsumerOfTheFinalPublicationPrimitive() {
    assertThrows(
        AssertionError.class,
        () ->
            ProductAssuranceArchitectureRules.publicationIsCentralizedBehindRequestPathAccess()
                .check(
                    new ClassFileImporter()
                        .importClasses(ArchitecturePublicationBypassFixture.class)));
  }
}
