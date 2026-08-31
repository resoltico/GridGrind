package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;

import com.tngtech.archunit.base.DescribedPredicate;
import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import com.tngtech.archunit.lang.CompositeArchRule;
import com.tngtech.archunit.lang.ConditionEvents;
import com.tngtech.archunit.lang.SimpleConditionEvent;
import java.util.List;

/** Defines bytecode rules for GridGrind's closed domain-union implementation shape. */
@SuppressWarnings("PMD.UseUtilityClass")
final class ProductDomainShapeArchitectureRules {
  private static final DescribedPredicate<JavaClass> SEALED_INTERFACES =
      new DescribedPredicate<>("sealed interfaces") {
        @Override
        public boolean test(JavaClass input) {
          return input.isInterface() && input.isSealed();
        }
      };

  private static final ArchCondition<JavaClass> USE_CLOSED_VARIANTS =
      new ArchCondition<>(
          "permit only records, enums, sealed subinterfaces, or exception implementations") {
        @Override
        public void check(JavaClass item, ConditionEvents events) {
          List<String> invalidVariants =
              item.getPermittedSubclasses().orElseThrow().stream()
                  .filter(variant -> !isAllowedVariant(variant))
                  .map(JavaClass::getName)
                  .sorted()
                  .toList();
          if (!invalidVariants.isEmpty()) {
            events.add(
                SimpleConditionEvent.violated(
                    item,
                    "Sealed interface <"
                        + item.getName()
                        + "> permits invalid variants "
                        + invalidVariants));
          }
        }
      };

  private static final ArchCondition<JavaClass> SEALED_CLASSES_ARE_EXCEPTION_BASES =
      new ArchCondition<>("use sealed classes only for typed exception bases") {
        @Override
        public void check(JavaClass item, ConditionEvents events) {
          if (item.isSealed() && !item.isInterface() && !item.isAssignableTo(Throwable.class)) {
            events.add(
                SimpleConditionEvent.violated(
                    item,
                    "Sealed class <"
                        + item.getName()
                        + "> is not assignable to <"
                        + Throwable.class.getName()
                        + ">"));
          }
        }
      };

  ProductDomainShapeArchitectureRules() {}

  static ArchRule sealedInterfacesUseClosedDomainVariants() {
    return classes()
        .that(SEALED_INTERFACES)
        .should(USE_CLOSED_VARIANTS)
        .because(
            "closed alternatives are records or finite enums and exception families remain typed");
  }

  static ArchRule sealedClassesModelOnlyExceptionBases() {
    return classes()
        .should(SEALED_CLASSES_ARE_EXCEPTION_BASES)
        .because("sealed class inheritance is reserved for typed exception bases");
  }

  static ArchRule sealedTypesUseClosedDomainShapes() {
    return CompositeArchRule.of(sealedInterfacesUseClosedDomainVariants())
        .and(sealedClassesModelOnlyExceptionBases())
        .as("sealed types preserve closed GridGrind domain shapes");
  }

  static ArchCondition<JavaClass> closedDomainVariantsAreRequired() {
    return USE_CLOSED_VARIANTS;
  }

  static boolean isAllowedVariant(JavaClass variant) {
    return variant.isRecord()
        || variant.isEnum()
        || (variant.isInterface() && variant.isSealed())
        || variant.isAssignableTo(Throwable.class);
  }
}
