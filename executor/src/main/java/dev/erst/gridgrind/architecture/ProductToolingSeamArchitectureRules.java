package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;

import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.domain.JavaCodeUnitAccess;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import com.tngtech.archunit.lang.ConditionEvents;
import com.tngtech.archunit.lang.SimpleConditionEvent;
import java.lang.invoke.MethodHandles;
import java.util.function.Predicate;

/** Defines bytecode access-site rules for GridGrind's centralized Apache POI tooling seams. */
@SuppressWarnings("PMD.UseUtilityClass")
final class ProductToolingSeamArchitectureRules {
  private static final String ENGINE_RUNTIME_PACKAGE = "dev.erst.gridgrind.engine.runtime..";
  private static final String EXCEL_PACKAGE = "dev.erst.gridgrind.excel..";
  private static final String FORMULA_WRITE_SUPPORT =
      "dev.erst.gridgrind.excel.ExcelFormulaWriteSupport";
  private static final String PRIVATE_ACCESS_SUPPORT =
      "dev.erst.gridgrind.excel.PoiPrivateAccessSupport";

  private static final ArchCondition<JavaClass> NOT_CALL_POI_FORMULA_WRITES =
      notAccessMethodsMatching(
          "access Apache POI formula-writing methods directly",
          ProductToolingSeamArchitectureRules::isPoiFormulaWrite);

  private static final ArchCondition<JavaClass> NOT_USE_PRIVATE_REFLECTION =
      notAccessMethodsMatching(
          "access private-reflection entry points directly",
          ProductToolingSeamArchitectureRules::isPrivateReflectionAccess);

  ProductToolingSeamArchitectureRules() {}

  static ArchRule formulaWritesStayCentralized() {
    return classes()
        .that()
        .resideInAnyPackage(ENGINE_RUNTIME_PACKAGE, EXCEL_PACKAGE)
        .and()
        .doNotHaveFullyQualifiedName(FORMULA_WRITE_SUPPORT)
        .should(NOT_CALL_POI_FORMULA_WRITES)
        .because("formula normalization and error translation require one write boundary");
  }

  static ArchRule privateReflectionStaysCentralized() {
    return classes()
        .that()
        .resideInAnyPackage(ENGINE_RUNTIME_PACKAGE, EXCEL_PACKAGE)
        .and()
        .doNotHaveFullyQualifiedName(PRIVATE_ACCESS_SUPPORT)
        .should(NOT_USE_PRIVATE_REFLECTION)
        .because("private dependency access requires one compatibility-tested boundary");
  }

  static ArchCondition<JavaClass> directPoiFormulaWritesAreForbidden() {
    return NOT_CALL_POI_FORMULA_WRITES;
  }

  static ArchCondition<JavaClass> privateReflectionIsForbidden() {
    return NOT_USE_PRIVATE_REFLECTION;
  }

  private static ArchCondition<JavaClass> notAccessMethodsMatching(
      String description, Predicate<JavaCodeUnitAccess<?>> forbiddenAccess) {
    return new ArchCondition<>("not " + description) {
      @Override
      public void check(JavaClass item, ConditionEvents events) {
        for (JavaCodeUnitAccess<?> access : item.getCodeUnitAccessesFromSelf()) {
          if (forbiddenAccess.test(access)) {
            events.add(SimpleConditionEvent.violated(access, access.getDescription()));
          }
        }
      }
    };
  }

  private static boolean isPoiFormulaWrite(JavaCodeUnitAccess<?> access) {
    return "setCellFormula".equals(access.getTarget().getName())
        && access.getTargetOwner().getPackageName().startsWith("org.apache.poi.");
  }

  private static boolean isPrivateReflectionAccess(JavaCodeUnitAccess<?> access) {
    String ownerName = access.getTargetOwner().getName();
    String methodName = access.getTarget().getName();
    return (MethodHandles.class.getName().equals(ownerName) && "privateLookupIn".equals(methodName))
        || (Class.class.getName().equals(ownerName) && methodName.startsWith("getDeclared"))
        || (ownerName.startsWith("java.lang.reflect.")
            && ("setAccessible".equals(methodName) || "trySetAccessible".equals(methodName)));
  }
}
