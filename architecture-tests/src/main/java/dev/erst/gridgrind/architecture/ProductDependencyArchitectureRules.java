package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;
import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.noClasses;

import com.tngtech.archunit.base.DescribedPredicate;
import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.domain.JavaModifier;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import com.tngtech.archunit.lang.ConditionEvents;
import com.tngtech.archunit.lang.SimpleConditionEvent;

/** Protects engine boundaries that module readability and exports alone cannot express. */
enum ProductDependencyArchitectureRules {
  ;
  private static final String ENGINE_API_PACKAGE = "dev.erst.gridgrind.engine.api";
  private static final String ENGINE_RUNTIME_PACKAGE = "dev.erst.gridgrind.engine.runtime..";
  private static final String[] NATIVE_PACKAGES = {
    "dev.erst.gridgrind.engine.backend..", "dev.erst.gridgrind.engine.poi..",
    "dev.erst.gridgrind.engine.cells..", "dev.erst.gridgrind.engine.styles..",
    "dev.erst.gridgrind.engine.formulas..", "dev.erst.gridgrind.engine.tables..",
    "dev.erst.gridgrind.engine.autofilter..", "dev.erst.gridgrind.engine.validation..",
    "dev.erst.gridgrind.engine.conditionalformatting..", "dev.erst.gridgrind.engine.charts..",
    "dev.erst.gridgrind.engine.pivots..", "dev.erst.gridgrind.engine.drawing..",
    "dev.erst.gridgrind.engine.comments..", "dev.erst.gridgrind.engine.links..",
    "dev.erst.gridgrind.engine.names..", "dev.erst.gridgrind.engine.print..",
    "dev.erst.gridgrind.engine.layout..", "dev.erst.gridgrind.engine.protection..",
    "dev.erst.gridgrind.engine.security..", "dev.erst.gridgrind.engine.customxml..",
    "dev.erst.gridgrind.engine.packaging.."
  };
  private static final String ENGINE_API_RUNTIME_BRIDGE_PATTERN =
      "^dev\\.erst\\.gridgrind\\.engine\\.api\\."
          + "(GridGrindEngine(\\$.*)?|GridGrindProblems|GridGrindRequestRequirements)$";

  private static final DescribedPredicate<JavaClass> ENGINE_INTERNAL =
      new DescribedPredicate<>("engine implementation types") {
        @Override
        public boolean test(JavaClass input) {
          return new EngineImplementationTypeClassifier().isImplementationType(input);
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

  static ArchRule coreDoesNotOwnConcreteHostTypes() {
    return classes()
        .that()
        .resideInAPackage("dev.erst.gridgrind.core..")
        .should()
        .haveNameNotMatching(
            ".*(Cli|Docker|Container|TerminalSession|TerminalCommand|TerminalWindow).*")
        .because("product discovery and HostContext are independent of concrete host transports");
  }

  static ArchRule coreDoesNotDependOnWorkbookOrJsonLibraries() {
    return noClasses()
        .that()
        .resideInAPackage("dev.erst.gridgrind.core..")
        .should()
        .dependOnClassesThat()
        .resideInAnyPackage("org.apache.poi..", "tools.jackson..", "com.fasterxml.jackson..")
        .because("the canonical model is independent of workbook and wire libraries");
  }

  static ArchRule workbookImplementationDoesNotDependOnExecutionRuntime() {
    return noClasses()
        .that()
        .resideInAnyPackage(NATIVE_PACKAGES)
        .should()
        .dependOnClassesThat()
        .resideInAPackage(ENGINE_RUNTIME_PACKAGE)
        .because("workbook mechanics sit below request execution and orchestration");
  }

  static ArchRule poiDependenciesStayInNativeAdapters() {
    return noClasses()
        .that()
        .resideInAPackage("dev.erst.gridgrind.engine..")
        .and()
        .resideOutsideOfPackages(NATIVE_PACKAGES)
        .should()
        .dependOnClassesThat()
        .resideInAnyPackage("org.apache.poi..", "org.openxmlformats..")
        .because(
            "runtime, publication and API use native operations without depending on POI types");
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
        .because("engine internals and POI types must not leak through the public module seam");
  }

  static ArchCondition<JavaClass> exportedApiMustNotExposeImplementationTypes() {
    return NOT_EXPOSE_ENGINE_IMPLEMENTATION_TYPES;
  }

  private static DescribedPredicate<JavaClass> exportedEngineApiClasses() {
    return new DescribedPredicate<>("effectively public classes in the exported engine API") {
      @Override
      public boolean test(JavaClass input) {
        return ENGINE_API_PACKAGE.equals(input.getPackageName())
            && allEnclosingClassesArePublic(input);
      }
    };
  }

  private static boolean allEnclosingClassesArePublic(JavaClass input) {
    JavaClass candidate = input;
    while (true) {
      if (!candidate.getModifiers().contains(JavaModifier.PUBLIC)) return false;
      var enclosing = candidate.getEnclosingClass();
      if (enclosing.isEmpty()) return true;
      candidate = enclosing.orElseThrow();
    }
  }
}
