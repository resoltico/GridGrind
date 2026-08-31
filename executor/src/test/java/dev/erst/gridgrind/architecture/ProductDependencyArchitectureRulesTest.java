package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.importer.ClassFileImporter;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import dev.erst.gridgrind.architecture.fixture.ArchitectureExcelLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureExtensiblePackageRuntimeLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureFinalProtectedRuntimeLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureInheritedRuntimeLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitecturePrivateRuntimeLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureProtectedRuntimeLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureRuntimeInterfaceLeakFixture;
import dev.erst.gridgrind.architecture.fixture.ArchitectureRuntimeLeakFixture;
import dev.erst.gridgrind.authoring.GridGrindPlan;
import dev.erst.gridgrind.cli.GridGrindCli;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies dependency, module-assignment, and exported-surface rule behavior. */
class ProductDependencyArchitectureRulesTest {
  private static final String PACKAGE_PRIVATE_OUTER_PUBLIC_NESTED =
      "dev.erst.gridgrind.architecture.fixture.ArchitectureEnclosingVisibilityFixture$PublicNested";

  @Test
  void mapsEveryProductPackageFamilyAndIgnoresVerificationFixtures() {
    assertEquals(Optional.of("excel-foundation"), moduleFor(ExcelComparisonOperator.class));
    assertEquals(Optional.of("contract"), moduleFor(WorkbookPlan.class));
    assertEquals(Optional.of("engine"), moduleFor(GridGrindEngine.class));
    assertEquals(Optional.of("engine"), moduleFor(dev.erst.gridgrind.excel.ExcelWorkbook.class));
    assertEquals(Optional.of("authoring-java"), moduleFor(GridGrindPlan.class));
    assertEquals(Optional.of("cli"), moduleFor(GridGrindCli.class));
    assertEquals(Optional.empty(), moduleFor(ArchitectureRuntimeLeakFixture.class));
  }

  @Test
  void reviewedRuntimeBridgeSetIsExact() {
    assertTrue(
        ProductDependencyArchitectureRules.isReviewedRuntimeBridge(
            imported(GridGrindEngine.class)));
    assertTrue(
        ProductDependencyArchitectureRules.isReviewedRuntimeBridge(
            imported(GridGrindProblems.class)));
    assertTrue(
        ProductDependencyArchitectureRules.isReviewedRuntimeBridge(
            imported(GridGrindRequestRequirements.class)));
    assertFalse(
        ProductDependencyArchitectureRules.isReviewedRuntimeBridge(
            imported(GridGrindRequestDoctor.class)));
  }

  @Test
  void exportedEngineApiPredicateRejectsPrivateNestedImplementationClasses() {
    var apiClasses = new ClassFileImporter().importPackages("dev.erst.gridgrind.engine.api");

    assertTrue(
        ProductDependencyArchitectureRules.exportedEngineApiClasses()
            .test(apiClasses.get(GridGrindEngine.class.getName())));
    assertFalse(
        ProductDependencyArchitectureRules.exportedEngineApiClasses()
            .test(
                apiClasses.get(
                    "dev.erst.gridgrind.engine.api.GridGrindEngine$ProductionRequestExecutor")));
  }

  @Test
  void enclosingClassVisibilityTreatsPublicAndPrivateNestedTypesPrecisely() {
    var apiClasses = new ClassFileImporter().importPackages("dev.erst.gridgrind.engine.api");

    assertTrue(
        ProductDependencyArchitectureRules.allEnclosingClassesArePublic(imported(Map.Entry.class)));
    assertTrue(
        ProductDependencyArchitectureRules.allEnclosingClassesArePublic(
            imported(GridGrindEngine.class)));
    assertFalse(
        ProductDependencyArchitectureRules.allEnclosingClassesArePublic(
            apiClasses.get(
                "dev.erst.gridgrind.engine.api.GridGrindEngine$ProductionRequestExecutor")));
    assertFalse(
        ProductDependencyArchitectureRules.allEnclosingClassesArePublic(
            new ClassFileImporter()
                .importPackages("dev.erst.gridgrind.architecture.fixture")
                .get(PACKAGE_PRIVATE_OUTER_PUBLIC_NESTED)));
  }

  @Test
  void implementationTypePredicateDistinguishesRuntimeWorkbookAndFoundationPackages() {
    EngineImplementationTypeClassifier classifier = new EngineImplementationTypeClassifier();

    assertTrue(
        classifier.isImplementationType(
            imported(dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor.class)));
    assertTrue(
        classifier.isImplementationType(imported(dev.erst.gridgrind.excel.ExcelWorkbook.class)));
    assertFalse(classifier.isImplementationType(imported(ExcelComparisonOperator.class)));
    assertTrue(
        EngineImplementationTypeClassifier.isInOrBelow(
            "dev.erst.gridgrind.engine.runtime", "dev.erst.gridgrind.engine.runtime"));
    assertFalse(
        EngineImplementationTypeClassifier.isInOrBelow(
            "dev.erst.gridgrind.engine.runtimex", "dev.erst.gridgrind.engine.runtime"));
    assertTrue(
        EngineImplementationTypeClassifier.isInOrBelow(
            "dev.erst.gridgrind.engine.runtime.detail", "dev.erst.gridgrind.engine.runtime"));
  }

  @Test
  void exportedApiConditionRejectsPublicAndGenericRuntimeTypeLeaks() {
    AssertionError violation =
        assertViolation(ArchitectureRuntimeLeakFixture.class, implementationLeakCondition());

    assertTrue(violation.getMessage().contains("GridGrindRequestDoctor"));
  }

  @Test
  void exportedApiConditionRejectsProtectedRuntimeTypeLeaksFromExtensibleClasses() {
    AssertionError violation =
        assertViolation(
            ArchitectureProtectedRuntimeLeakFixture.class, implementationLeakCondition());

    assertTrue(violation.getMessage().contains("ArchitectureProtectedRuntimeLeakFixture"));
  }

  @Test
  void exportedApiConditionRejectsRuntimeSupertypesAndInheritedMembers() {
    AssertionError interfaceViolation =
        assertViolation(
            ArchitectureRuntimeInterfaceLeakFixture.class, implementationLeakCondition());
    AssertionError inheritedMemberViolation =
        assertViolation(
            ArchitectureInheritedRuntimeLeakFixture.class, implementationLeakCondition());

    assertTrue(interfaceViolation.getMessage().contains("GridGrindRequestExecutor"));
    assertTrue(inheritedMemberViolation.getMessage().contains("GridGrindRequestDoctor"));
  }

  @Test
  void exportedApiConditionRejectsWorkbookImplementationTypeLeaks() {
    AssertionError violation =
        assertViolation(ArchitectureExcelLeakFixture.class, implementationLeakCondition());

    assertTrue(violation.getMessage().contains("ExcelWorkbook"));
  }

  @Test
  void exportedApiConditionAllowsPrivateImplementationTypeUsage() {
    ArchRule deliberateCheck = classes().should(implementationLeakCondition());

    org.junit.jupiter.api.Assertions.assertDoesNotThrow(
        () ->
            deliberateCheck.check(
                new ClassFileImporter()
                    .importClasses(ArchitecturePrivateRuntimeLeakFixture.class)));
  }

  @Test
  void implementationTypeExtractionRetainsEveryExternallyReachableLeakChannel() {
    ExportedApiImplementationTypes implementationTypes = new ExportedApiImplementationTypes();

    assertTrue(
        implementationTypes
            .namesFor(imported(ArchitectureRuntimeInterfaceLeakFixture.class))
            .contains("dev.erst.gridgrind.engine.runtime.GridGrindRequestExecutor"));
    assertTrue(
        implementationTypes
            .namesFor(imported(ArchitectureInheritedRuntimeLeakFixture.class))
            .contains("dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor"));
    assertTrue(
        implementationTypes
            .namesFor(imported(ArchitectureExcelLeakFixture.class))
            .contains("dev.erst.gridgrind.excel.ExcelWorkbook"));
  }

  @Test
  void externallyReachableMemberClassificationRespectsVisibilityAndExtensibility() {
    ExportedApiImplementationTypes implementationTypes = new ExportedApiImplementationTypes();

    assertTrue(
        implementationTypes.isExternallyReachable(
            imported(ArchitectureRuntimeLeakFixture.class),
            member(ArchitectureRuntimeLeakFixture.class, "leakedRuntimeType")));
    assertFalse(
        implementationTypes.isExternallyReachable(
            imported(ArchitecturePrivateRuntimeLeakFixture.class),
            member(ArchitecturePrivateRuntimeLeakFixture.class, "internalRuntimeType")));
    assertTrue(
        implementationTypes.isExternallyReachable(
            imported(ArchitectureProtectedRuntimeLeakFixture.class),
            member(ArchitectureProtectedRuntimeLeakFixture.class, "leakedRuntimeType")));
    assertFalse(
        implementationTypes.isExternallyReachable(
            imported(ArchitectureFinalProtectedRuntimeLeakFixture.class),
            member(ArchitectureFinalProtectedRuntimeLeakFixture.class, "internalRuntimeType")));
    assertFalse(
        implementationTypes.isExternallyReachable(
            imported(ArchitectureExtensiblePackageRuntimeLeakFixture.class),
            member(ArchitectureExtensiblePackageRuntimeLeakFixture.class, "internalRuntimeType")));
  }

  @Test
  void productModuleSliceAssignmentMapsKnownPackagesAndRejectsVerificationPackages() {
    assertEquals(
        "GridGrind production modules",
        ProductDependencyArchitectureRules.PRODUCT_MODULE_ASSIGNMENT.getDescription());
    assertEquals(
        com.tngtech.archunit.library.dependencies.SliceIdentifier.of("engine"),
        ProductDependencyArchitectureRules.PRODUCT_MODULE_ASSIGNMENT.getIdentifierOf(
            imported(dev.erst.gridgrind.excel.ExcelWorkbook.class)));
    assertEquals(
        com.tngtech.archunit.library.dependencies.SliceIdentifier.ignore(),
        ProductDependencyArchitectureRules.PRODUCT_MODULE_ASSIGNMENT.getIdentifierOf(
            imported(ArchitectureRuntimeLeakFixture.class)));
  }

  private static Optional<String> moduleFor(Class<?> type) {
    return ProductDependencyArchitectureRules.productModuleFor(imported(type));
  }

  private static JavaClass imported(Class<?> type) {
    return new ClassFileImporter().importClasses(type).get(type.getName());
  }

  private static com.tngtech.archunit.core.domain.JavaMember member(
      Class<?> type, String memberName) {
    return imported(type).getMembers().stream()
        .filter(member -> member.getName().equals(memberName))
        .findFirst()
        .orElseThrow();
  }

  private static ArchCondition<JavaClass> implementationLeakCondition() {
    return ProductDependencyArchitectureRules.exportedApiMustNotExposeImplementationTypes();
  }

  private static AssertionError assertViolation(
      Class<?> violationFixture, ArchCondition<JavaClass> condition) {
    ArchRule deliberateViolation = classes().should(condition);
    return assertThrows(
        AssertionError.class,
        () -> deliberateViolation.check(new ClassFileImporter().importClasses(violationFixture)));
  }
}
