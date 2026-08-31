package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.importer.ClassFileImporter;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import dev.erst.gridgrind.architecture.fixture.ArchitectureBenignReflectionMethodFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureMethodReferenceViolationFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureNonPoiFormulaMethodFixture;
import dev.erst.gridgrind.excel.ExcelFormulaWriteSupport;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

/** Verifies direct and method-reference access to centralized tooling seams. */
class ProductToolingSeamArchitectureRulesTest {
  @Test
  void directPoiFormulaWriteConditionRejectsTheCentralizedWriterWhenAppliedWithoutItsExemption() {
    AssertionError violation =
        assertViolation(
            ExcelFormulaWriteSupport.class,
            ProductToolingSeamArchitectureRules.directPoiFormulaWritesAreForbidden());

    assertTrue(violation.getMessage().contains("setCellFormula"));
  }

  @Test
  void privateReflectionConditionRejectsTheCentralizedBridgeWhenAppliedWithoutItsExemption() {
    var engineWorkbookClasses = new ClassFileImporter().importPackages("dev.erst.gridgrind.excel");
    ArchRule deliberateViolation =
        classes()
            .that()
            .haveFullyQualifiedName("dev.erst.gridgrind.excel.PoiPrivateAccessSupport")
            .should(ProductToolingSeamArchitectureRules.privateReflectionIsForbidden());

    AssertionError violation =
        assertThrows(AssertionError.class, () -> deliberateViolation.check(engineWorkbookClasses));

    assertTrue(violation.getMessage().contains("privateLookupIn"));
  }

  @Test
  void formulaWriteConditionRejectsMethodReferences() {
    AssertionError violation =
        assertViolation(
            ArchitectureMethodReferenceViolationFixture.class,
            ProductToolingSeamArchitectureRules.directPoiFormulaWritesAreForbidden());

    assertTrue(violation.getMessage().contains("setCellFormula"));
  }

  @Test
  void privateReflectionConditionRejectsMethodReferences() {
    AssertionError violation =
        assertViolation(
            ArchitectureMethodReferenceViolationFixture.class,
            ProductToolingSeamArchitectureRules.privateReflectionIsForbidden());

    assertTrue(violation.getMessage().contains("trySetAccessible"));
    assertTrue(violation.getMessage().contains("setAccessible"));
  }

  @Test
  void privateReflectionConditionRejectsDeclaredMetadataMethodReferences() {
    AssertionError violation =
        assertViolation(
            ArchitectureMethodReferenceViolationFixture.class,
            ProductToolingSeamArchitectureRules.privateReflectionIsForbidden());

    assertTrue(violation.getMessage().contains("getDeclaredFields"));
  }

  @Test
  void accessConditionsAllowBenignSameNameAndReflectionMetadataMethods() {
    Assertions.assertDoesNotThrow(
        () ->
            classes()
                .should(ProductToolingSeamArchitectureRules.directPoiFormulaWritesAreForbidden())
                .check(
                    new ClassFileImporter()
                        .importClasses(ArchitectureNonPoiFormulaMethodFixture.class)));
    Assertions.assertDoesNotThrow(
        () ->
            classes()
                .should(ProductToolingSeamArchitectureRules.privateReflectionIsForbidden())
                .check(
                    new ClassFileImporter()
                        .importClasses(ArchitectureBenignReflectionMethodFixture.class)));
  }

  private static AssertionError assertViolation(
      Class<?> violationFixture, ArchCondition<JavaClass> condition) {
    ArchRule deliberateViolation = classes().should(condition);
    return assertThrows(
        AssertionError.class,
        () -> deliberateViolation.check(new ClassFileImporter().importClasses(violationFixture)));
  }
}
