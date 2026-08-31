package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.base.DescribedPredicate.not;
import static com.tngtech.archunit.core.domain.JavaClass.Predicates.resideInAPackage;
import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;
import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.noClasses;
import static com.tngtech.archunit.library.dependencies.SlicesRuleDefinition.slices;

import com.tngtech.archunit.base.DescribedPredicate;
import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.domain.JavaModifier;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import com.tngtech.archunit.lang.ConditionEvents;
import com.tngtech.archunit.lang.SimpleConditionEvent;
import com.tngtech.archunit.library.dependencies.SliceAssignment;
import com.tngtech.archunit.library.dependencies.SliceIdentifier;
import java.util.Optional;

/** Defines GridGrind's production module, package, and exported-surface architecture rules. */
@SuppressWarnings("PMD.UseUtilityClass")
final class ProductDependencyArchitectureRules {
  private static final String AUTHORING_PACKAGE = "dev.erst.gridgrind.authoring..";
  private static final String CLI_PACKAGE = "dev.erst.gridgrind.cli..";
  private static final String CONTRACT_PACKAGE = "dev.erst.gridgrind.contract..";
  private static final String ENGINE_API_PACKAGE = "dev.erst.gridgrind.engine.api";
  private static final String ENGINE_PACKAGE = "dev.erst.gridgrind.engine..";
  private static final String ENGINE_RUNTIME_PACKAGE = "dev.erst.gridgrind.engine.runtime..";
  private static final String EXCEL_PACKAGE = "dev.erst.gridgrind.excel..";
  private static final String EXCEL_FOUNDATION_PACKAGE = "dev.erst.gridgrind.excel.foundation..";
  private static final String ENGINE_API_RUNTIME_BRIDGE_PATTERN =
      "^dev\\.erst\\.gridgrind\\.engine\\.api\\."
          + "(GridGrindEngine(\\$.*)?|GridGrindProblems|GridGrindRequestRequirements)$";

  private static final DescribedPredicate<JavaClass> ENGINE_NON_FOUNDATION =
      resideInAPackage(ENGINE_PACKAGE)
          .or(resideInAPackage(EXCEL_PACKAGE).and(not(resideInAPackage(EXCEL_FOUNDATION_PACKAGE))));

  private static final DescribedPredicate<JavaClass> ENGINE_INTERNAL =
      resideInAPackage(ENGINE_RUNTIME_PACKAGE)
          .or(resideInAPackage(EXCEL_PACKAGE).and(not(resideInAPackage(EXCEL_FOUNDATION_PACKAGE))));

  static final SliceAssignment PRODUCT_MODULE_ASSIGNMENT =
      new SliceAssignment() {
        @Override
        public String getDescription() {
          return "GridGrind production modules";
        }

        @Override
        public SliceIdentifier getIdentifierOf(JavaClass javaClass) {
          return productModuleFor(javaClass)
              .map(SliceIdentifier::of)
              .orElseGet(SliceIdentifier::ignore);
        }
      };

  private static final ArchCondition<JavaClass> NOT_EXPOSE_ENGINE_IMPLEMENTATION_TYPES =
      new ArchCondition<>("not expose engine implementation types through the exported API") {
        @Override
        public void check(JavaClass item, ConditionEvents events) {
          var implementationTypes = new ExportedApiImplementationTypes().namesFor(item);
          if (!implementationTypes.isEmpty()) {
            events.add(
                SimpleConditionEvent.violated(
                    item,
                    "Exported API type <"
                        + item.getName()
                        + "> exposes engine implementation types "
                        + implementationTypes));
          }
        }
      };

  ProductDependencyArchitectureRules() {}

  static ArchRule productModulesAreFreeOfCycles() {
    return slices()
        .assignedFrom(PRODUCT_MODULE_ASSIGNMENT)
        .namingSlices("$1")
        .should()
        .beFreeOfCycles()
        .because("the published product dependency graph is acyclic");
  }

  static ArchRule productClassesBelongToKnownModules() {
    return classes()
        .should()
        .resideInAnyPackage(
            AUTHORING_PACKAGE, CLI_PACKAGE, CONTRACT_PACKAGE, ENGINE_PACKAGE, EXCEL_PACKAGE)
        .because("each imported production class belongs to one declared GridGrind package family");
  }

  static ArchRule excelFoundationDoesNotDependOnHigherProductModules() {
    DescribedPredicate<JavaClass> higherProductClass =
        resideInAPackage("dev.erst.gridgrind..")
            .and(not(resideInAPackage(EXCEL_FOUNDATION_PACKAGE)));
    return noClasses()
        .that()
        .resideInAPackage(EXCEL_FOUNDATION_PACKAGE)
        .should()
        .dependOnClassesThat(higherProductClass)
        .because("the shared Excel vocabulary is the bottom of the product graph");
  }

  static ArchRule contractDependsOnlyOnExcelFoundation() {
    return noClasses()
        .that()
        .resideInAPackage(CONTRACT_PACKAGE)
        .should()
        .dependOnClassesThat(ENGINE_NON_FOUNDATION)
        .because("the contract may share only the independent Excel foundation vocabulary");
  }

  static ArchRule authoringStaysOnContractBoundary() {
    DescribedPredicate<JavaClass> forbiddenDependency =
        ENGINE_NON_FOUNDATION.or(resideInAPackage(CLI_PACKAGE));
    return noClasses()
        .that()
        .resideInAPackage(AUTHORING_PACKAGE)
        .should()
        .dependOnClassesThat(forbiddenDependency)
        .because("the Java authoring API is a contract-only downstream surface");
  }

  static ArchRule cliUsesOnlyExportedEngineApi() {
    DescribedPredicate<JavaClass> forbiddenDependency =
        ENGINE_INTERNAL.or(resideInAPackage(AUTHORING_PACKAGE));
    return noClasses()
        .that()
        .resideInAPackage(CLI_PACKAGE)
        .should()
        .dependOnClassesThat(forbiddenDependency)
        .because("the CLI must consume only the exported engine API and shared foundation");
  }

  static ArchRule engineDoesNotDependOnDownstreamAdapters() {
    DescribedPredicate<JavaClass> downstreamAdapter =
        resideInAPackage(AUTHORING_PACKAGE).or(resideInAPackage(CLI_PACKAGE));
    return noClasses()
        .that(ENGINE_NON_FOUNDATION)
        .should()
        .dependOnClassesThat(downstreamAdapter)
        .because(
            "the engine must remain independent of downstream authoring and transport adapters");
  }

  static ArchRule workbookImplementationDoesNotDependOnExecutionRuntime() {
    return noClasses()
        .that()
        .resideInAPackage(EXCEL_PACKAGE)
        .should()
        .dependOnClassesThat()
        .resideInAPackage(ENGINE_RUNTIME_PACKAGE)
        .because("workbook mechanics sit below request execution and orchestration");
  }

  static ArchRule engineApiRuntimeDependenciesStayInOwnedBridges() {
    return noClasses()
        .that()
        .resideInAPackage(ENGINE_API_PACKAGE)
        .and()
        .haveNameNotMatching(ENGINE_API_RUNTIME_BRIDGE_PATTERN)
        .should()
        .dependOnClassesThat(ENGINE_INTERNAL)
        .because(
            "the exported API exposes contracts while named facades own implementation wiring");
  }

  static ArchRule engineApiDoesNotExposeImplementationTypes() {
    return classes()
        .that(exportedEngineApiClasses())
        .should(NOT_EXPOSE_ENGINE_IMPLEMENTATION_TYPES)
        .because(
            "unexported engine implementation types must not leak through the public module seam");
  }

  static ArchCondition<JavaClass> exportedApiMustNotExposeImplementationTypes() {
    return NOT_EXPOSE_ENGINE_IMPLEMENTATION_TYPES;
  }

  static Optional<String> productModuleFor(JavaClass javaClass) {
    String packageName = javaClass.getPackageName();
    if (packageName.startsWith("dev.erst.gridgrind.excel.foundation")) {
      return Optional.of("excel-foundation");
    }
    if (packageName.startsWith("dev.erst.gridgrind.contract")) {
      return Optional.of("contract");
    }
    if (packageName.startsWith("dev.erst.gridgrind.engine")
        || packageName.startsWith("dev.erst.gridgrind.excel")) {
      return Optional.of("engine");
    }
    if (packageName.startsWith("dev.erst.gridgrind.authoring")) {
      return Optional.of("authoring-java");
    }
    if (packageName.startsWith("dev.erst.gridgrind.cli")) {
      return Optional.of("cli");
    }
    return Optional.empty();
  }

  static boolean isReviewedRuntimeBridge(JavaClass javaClass) {
    return javaClass.getName().matches(ENGINE_API_RUNTIME_BRIDGE_PATTERN);
  }

  static DescribedPredicate<JavaClass> exportedEngineApiClasses() {
    return new DescribedPredicate<>("effectively public classes in the exported engine API") {
      @Override
      public boolean test(JavaClass input) {
        return ENGINE_API_PACKAGE.equals(input.getPackageName())
            && allEnclosingClassesArePublic(input);
      }
    };
  }

  static boolean allEnclosingClassesArePublic(JavaClass input) {
    JavaClass candidate = input;
    while (true) {
      if (!candidate.getModifiers().contains(JavaModifier.PUBLIC)) {
        return false;
      }
      var enclosingClass = candidate.getEnclosingClass();
      if (enclosingClass.isEmpty()) {
        return true;
      }
      candidate = enclosingClass.orElseThrow();
    }
  }
}
