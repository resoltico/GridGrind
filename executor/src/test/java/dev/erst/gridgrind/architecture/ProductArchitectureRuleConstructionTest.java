package dev.erst.gridgrind.architecture;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import org.junit.jupiter.api.Test;

/** Verifies that every rule and custom condition can be constructed directly. */
class ProductArchitectureRuleConstructionTest {
  @Test
  void constructsEveryMandatoryRuleAndCustomCondition() {
    assertAll(
        () -> assertNotNull(new ProductArchitectureRules()),
        () -> assertNotNull(new EngineImplementationTypeClassifier()),
        () -> assertNotNull(new ExportedApiImplementationTypes()),
        () -> assertNotNull(new ProductDependencyArchitectureRules()),
        () -> assertNotNull(new ProductDomainShapeArchitectureRules()),
        () -> assertNotNull(new ProductToolingSeamArchitectureRules()),
        () -> assertNotNull(ProductDependencyArchitectureRules.productModulesAreFreeOfCycles()),
        () ->
            assertNotNull(ProductDependencyArchitectureRules.productClassesBelongToKnownModules()),
        () ->
            assertNotNull(
                ProductDependencyArchitectureRules
                    .excelFoundationDoesNotDependOnHigherProductModules()),
        () ->
            assertNotNull(
                ProductDependencyArchitectureRules.contractDependsOnlyOnExcelFoundation()),
        () -> assertNotNull(ProductDependencyArchitectureRules.authoringStaysOnContractBoundary()),
        () -> assertNotNull(ProductDependencyArchitectureRules.cliUsesOnlyExportedEngineApi()),
        () ->
            assertNotNull(
                ProductDependencyArchitectureRules.engineDoesNotDependOnDownstreamAdapters()),
        () ->
            assertNotNull(
                ProductDependencyArchitectureRules
                    .workbookImplementationDoesNotDependOnExecutionRuntime()),
        () ->
            assertNotNull(
                ProductDependencyArchitectureRules
                    .engineApiRuntimeDependenciesStayInOwnedBridges()),
        () ->
            assertNotNull(
                ProductDependencyArchitectureRules.engineApiDoesNotExposeImplementationTypes()),
        () ->
            assertNotNull(
                ProductDependencyArchitectureRules.exportedApiMustNotExposeImplementationTypes()),
        () ->
            assertNotNull(
                ProductDomainShapeArchitectureRules.sealedInterfacesUseClosedDomainVariants()),
        () ->
            assertNotNull(
                ProductDomainShapeArchitectureRules.sealedClassesModelOnlyExceptionBases()),
        () -> assertNotNull(ProductDomainShapeArchitectureRules.sealedTypesUseClosedDomainShapes()),
        () -> assertNotNull(ProductDomainShapeArchitectureRules.closedDomainVariantsAreRequired()),
        () -> assertNotNull(ProductToolingSeamArchitectureRules.formulaWritesStayCentralized()),
        () ->
            assertNotNull(ProductToolingSeamArchitectureRules.privateReflectionStaysCentralized()),
        () ->
            assertNotNull(ProductToolingSeamArchitectureRules.directPoiFormulaWritesAreForbidden()),
        () -> assertNotNull(ProductToolingSeamArchitectureRules.privateReflectionIsForbidden()));
  }
}
