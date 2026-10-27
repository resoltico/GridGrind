package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;
import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.noClasses;

import com.tngtech.archunit.base.DescribedPredicate;
import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.junit.ArchTest;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import com.tngtech.archunit.lang.conditions.ArchConditions;

/** Protects the single admitted executor, canonical semantics, and publication boundaries. */
final class ProductAssuranceArchitectureRules {
  private static final String CATALOG_FACTORY =
      "dev.erst.gridgrind.core.catalog.CatalogTypeEntryFactory";
  private static final String CANONICAL_OPERATION_CONTRACTS =
      "dev.erst.gridgrind.core.semantics.WorkbookOperationContracts";
  private static final String REQUEST_EXECUTOR =
      "dev.erst.gridgrind.engine.runtime.GridGrindRequestExecutor";
  private static final String ADMITTED_EXECUTOR =
      "dev.erst.gridgrind.engine.runtime.DefaultGridGrindRequestExecutor";
  private static final String ENGINE_RUNTIME_PACKAGE = "dev.erst.gridgrind.engine.runtime..";
  private static final String REQUEST_PATH_ACCESS =
      "dev.erst.gridgrind.engine.runtime.RequestPathAccess";
  private static final String REQUEST_PATH_PUBLICATION =
      "dev.erst.gridgrind.engine.runtime.RequestPathPublication";
  private static final String PUBLICATION_FILE_SUPPORT =
      "dev.erst.gridgrind.engine.runtime.PublicationFileOperations";

  private static final ArchCondition<JavaClass> USE_CANONICAL_OPERATION_CONTRACTS =
      ArchConditions.dependOnClassesThat(
          new DescribedPredicate<JavaClass>("canonical operation contracts") {
            @Override
            public boolean test(JavaClass input) {
              return CANONICAL_OPERATION_CONTRACTS.equals(input.getName());
            }
          });

  private ProductAssuranceArchitectureRules() {}

  @ArchTest
  static final ArchRule EXECUTION_USES_ONE_ADMITTED_EXECUTOR = executionUsesOneAdmittedExecutor();

  @ArchTest
  static final ArchRule CATALOG_USES_CANONICAL_OPERATION_CONTRACTS =
      catalogUsesCanonicalOperationContracts();

  @ArchTest
  static final ArchRule PUBLICATION_IS_CENTRALIZED_BEHIND_REQUEST_PATH_ACCESS =
      publicationIsCentralizedBehindRequestPathAccess();

  static ArchRule executionUsesOneAdmittedExecutor() {
    return classes()
        .that()
        .implement(REQUEST_EXECUTOR)
        .should()
        .haveFullyQualifiedName(ADMITTED_EXECUTOR)
        .because("all executable requests must cross shared admission before workbook mutation");
  }

  static ArchRule catalogUsesCanonicalOperationContracts() {
    return classes()
        .that()
        .haveFullyQualifiedName(CATALOG_FACTORY)
        .should(USE_CANONICAL_OPERATION_CONTRACTS)
        .because("catalog projection must share the operation-semantics owner used by admission");
  }

  static ArchCondition<JavaClass> canonicalOperationContractsAreRequired() {
    return USE_CANONICAL_OPERATION_CONTRACTS;
  }

  static ArchRule publicationIsCentralizedBehindRequestPathAccess() {
    return noClasses()
        .that()
        .resideInAPackage(ENGINE_RUNTIME_PACKAGE)
        .and()
        .doNotHaveFullyQualifiedName(REQUEST_PATH_ACCESS)
        .and()
        .doNotHaveFullyQualifiedName(REQUEST_PATH_PUBLICATION)
        .and()
        .doNotHaveFullyQualifiedName(PUBLICATION_FILE_SUPPORT)
        .should()
        .dependOnClassesThat()
        .haveFullyQualifiedName(REQUEST_PATH_PUBLICATION)
        .because("one descriptor-bound access capability owns final workbook publication");
  }
}
