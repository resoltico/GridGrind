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
  private static final String EXCEL_PACKAGE = "dev.erst.gridgrind.excel..";
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
