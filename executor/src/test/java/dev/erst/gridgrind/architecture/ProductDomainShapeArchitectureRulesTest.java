package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.importer.ClassFileImporter;
import com.tngtech.archunit.lang.ArchRule;
import dev.erst.gridgrind.architecture.fixture.ArchitectureRuntimeLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureSealedClassViolation;
import dev.erst.gridgrind.architecture.fixture.ArchitectureSealedUnionVariants;
import dev.erst.gridgrind.architecture.fixture.ArchitectureSealedUnionViolation;
import org.junit.jupiter.api.Test;

/** Verifies the closed-domain shape conditions and empty-population-safe sealed-class rule. */
class ProductDomainShapeArchitectureRulesTest {
  @Test
  void allowsEverySupportedSealedInterfaceVariantShape() {
    assertTrue(allowed(ArchitectureSealedUnionVariants.EnumVariant.class));
    assertTrue(allowed(ArchitectureSealedUnionVariants.RecordVariant.class));
    assertTrue(allowed(ArchitectureSealedUnionVariants.SealedSubinterface.class));
    assertTrue(allowed(ArchitectureSealedUnionVariants.TypedExceptionVariant.class));
    assertFalse(allowed(ArchitectureSealedUnionViolation.NonSealedVariant.class));
    assertFalse(allowed(ArchitectureSealedUnionViolation.OrdinaryVariant.class));
    assertFalse(allowed(ArchitectureSealedClassViolation.class));
  }

  @Test
  void sealedUnionConditionRejectsADeliberateOrdinaryClassVariant() {
    ArchRule deliberateViolation =
        classes().should(ProductDomainShapeArchitectureRules.closedDomainVariantsAreRequired());

    AssertionError violation =
        assertThrows(
            AssertionError.class,
            () ->
                deliberateViolation.check(
                    new ClassFileImporter().importClasses(ArchitectureSealedUnionViolation.class)));

    assertTrue(violation.getMessage().contains("OrdinaryVariant"));
  }

  @Test
  void sealedClassRuleRejectsADeliberateNonExceptionBase() {
    AssertionError violation =
        assertThrows(
            AssertionError.class,
            () ->
                ProductDomainShapeArchitectureRules.sealedClassesModelOnlyExceptionBases()
                    .check(
                        new ClassFileImporter()
                            .importClasses(ArchitectureSealedClassViolation.class)));

    assertTrue(violation.getMessage().contains("ArchitectureSealedClassViolation"));
  }

  @Test
  void sealedClassRuleAllowsAnOrdinaryNonSealedClass() {
    assertDoesNotThrow(
        () ->
            ProductDomainShapeArchitectureRules.sealedClassesModelOnlyExceptionBases()
                .check(
                    new ClassFileImporter().importClasses(ArchitectureRuntimeLeakFixture.class)));
  }

  @Test
  void sealedInterfaceRuleAcceptsValidUnionsAndIgnoresOrdinaryClassesInTheSameImport() {
    assertDoesNotThrow(
        () ->
            ProductDomainShapeArchitectureRules.sealedInterfacesUseClosedDomainVariants()
                .check(
                    new ClassFileImporter()
                        .importClasses(
                            ArchitectureSealedUnionVariants.class,
                            ArchitectureRuntimeLeakFixture.class,
                            ArchitectureSealedClassViolation.class)));
  }

  private static boolean allowed(Class<?> variant) {
    JavaClass imported = new ClassFileImporter().importClasses(variant).get(variant.getName());
    return ProductDomainShapeArchitectureRules.isAllowedVariant(imported);
  }
}
