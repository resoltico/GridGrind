package dev.erst.gridgrind.architecture;

import com.tngtech.archunit.junit.ArchTest;
import com.tngtech.archunit.junit.ArchTests;
import com.tngtech.archunit.lang.ArchRule;

/** Declares the complete mandatory GridGrind product architecture rule inventory. */
public enum ProductArchitectureRules {
  ;

  @ArchTest
  static final ArchRule CORE_DOES_NOT_DEPEND_ON_WORKBOOK_OR_JSON_LIBRARIES =
      ProductDependencyArchitectureRules.coreDoesNotDependOnWorkbookOrJsonLibraries();

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
  static final ArchTests TOOLING_SEAMS = ArchTests.in(ProductToolingSeamArchitectureRules.class);

  @ArchTest
  static final ArchRule SEALED_TYPES_USE_CLOSED_DOMAIN_SHAPES =
      ProductDomainShapeArchitectureRules.sealedTypesUseClosedDomainShapes();

  @ArchTest
  static final ArchTests ASSURANCE_BOUNDARIES =
      ArchTests.in(ProductAssuranceArchitectureRules.class);
}
