package dev.erst.gridgrind.architecture;

import com.tngtech.archunit.junit.ArchTest;
import com.tngtech.archunit.lang.ArchRule;
import java.util.List;

/** Declares the complete mandatory GridGrind product architecture rule inventory. */
@SuppressWarnings("PMD.UseUtilityClass")
public final class ProductArchitectureRules {
  @ArchTest
  static final ArchRule PRODUCT_MODULES_ARE_FREE_OF_CYCLES =
      ProductDependencyArchitectureRules.productModulesAreFreeOfCycles();

  @ArchTest
  static final ArchRule PRODUCT_CLASSES_BELONG_TO_KNOWN_MODULES =
      ProductDependencyArchitectureRules.productClassesBelongToKnownModules();

  @ArchTest
  static final ArchRule EXCEL_FOUNDATION_DOES_NOT_DEPEND_ON_HIGHER_PRODUCT_MODULES =
      ProductDependencyArchitectureRules.excelFoundationDoesNotDependOnHigherProductModules();

  @ArchTest
  static final ArchRule CONTRACT_DEPENDS_ONLY_ON_EXCEL_FOUNDATION =
      ProductDependencyArchitectureRules.contractDependsOnlyOnExcelFoundation();

  @ArchTest
  static final ArchRule AUTHORING_STAYS_ON_CONTRACT_BOUNDARY =
      ProductDependencyArchitectureRules.authoringStaysOnContractBoundary();

  @ArchTest
  static final ArchRule CLI_USES_ONLY_EXPORTED_ENGINE_API =
      ProductDependencyArchitectureRules.cliUsesOnlyExportedEngineApi();

  @ArchTest
  static final ArchRule ENGINE_DOES_NOT_DEPEND_ON_DOWNSTREAM_ADAPTERS =
      ProductDependencyArchitectureRules.engineDoesNotDependOnDownstreamAdapters();

  @ArchTest
  static final ArchRule WORKBOOK_IMPLEMENTATION_DOES_NOT_DEPEND_ON_EXECUTION_RUNTIME =
      ProductDependencyArchitectureRules.workbookImplementationDoesNotDependOnExecutionRuntime();

  @ArchTest
  static final ArchRule ENGINE_API_RUNTIME_DEPENDENCIES_STAY_IN_OWNED_BRIDGES =
      ProductDependencyArchitectureRules.engineApiRuntimeDependenciesStayInOwnedBridges();

  @ArchTest
  static final ArchRule ENGINE_API_DOES_NOT_EXPOSE_IMPLEMENTATION_TYPES =
      ProductDependencyArchitectureRules.engineApiDoesNotExposeImplementationTypes();

  @ArchTest
  static final ArchRule FORMULA_WRITES_STAY_CENTRALIZED =
      ProductToolingSeamArchitectureRules.formulaWritesStayCentralized();

  @ArchTest
  static final ArchRule PRIVATE_REFLECTION_STAYS_CENTRALIZED =
      ProductToolingSeamArchitectureRules.privateReflectionStaysCentralized();

  @ArchTest
  static final ArchRule SEALED_TYPES_USE_CLOSED_DOMAIN_SHAPES =
      ProductDomainShapeArchitectureRules.sealedTypesUseClosedDomainShapes();

  ProductArchitectureRules() {}

  static List<ArchRule> mandatoryRules() {
    return List.of(
        PRODUCT_MODULES_ARE_FREE_OF_CYCLES,
        PRODUCT_CLASSES_BELONG_TO_KNOWN_MODULES,
        EXCEL_FOUNDATION_DOES_NOT_DEPEND_ON_HIGHER_PRODUCT_MODULES,
        CONTRACT_DEPENDS_ONLY_ON_EXCEL_FOUNDATION,
        AUTHORING_STAYS_ON_CONTRACT_BOUNDARY,
        CLI_USES_ONLY_EXPORTED_ENGINE_API,
        ENGINE_DOES_NOT_DEPEND_ON_DOWNSTREAM_ADAPTERS,
        WORKBOOK_IMPLEMENTATION_DOES_NOT_DEPEND_ON_EXECUTION_RUNTIME,
        ENGINE_API_RUNTIME_DEPENDENCIES_STAY_IN_OWNED_BRIDGES,
        ENGINE_API_DOES_NOT_EXPOSE_IMPLEMENTATION_TYPES,
        FORMULA_WRITES_STAY_CENTRALIZED,
        PRIVATE_REFLECTION_STAYS_CENTRALIZED,
        SEALED_TYPES_USE_CLOSED_DOMAIN_SHAPES);
  }
}
