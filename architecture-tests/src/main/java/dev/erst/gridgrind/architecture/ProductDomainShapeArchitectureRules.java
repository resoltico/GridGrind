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
import java.util.stream.Stream;

/** Defines bytecode rules for GridGrind's closed domain-union implementation shape. */
enum ProductDomainShapeArchitectureRules {
  ;
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

  private static final ArchCondition<JavaClass> USE_RECORD_WIRE_VARIANTS =
      new ArchCondition<>("use records for every concrete wire variant") {
        @Override
        public void check(JavaClass item, ConditionEvents events) {
          wireVariants(item)
              .filter(variant -> !variant.isRecord())
              .forEach(
                  variant ->
                      events.add(
                          SimpleConditionEvent.violated(
                              item,
                              "Wire union <"
                                  + item.getName()
                                  + "> has non-record variant <"
                                  + variant.getName()
                                  + ">")));
        }
      };

  static ArchRule wireUnionsUseRecordVariants() {
    return classes()
        .that(
            new DescribedPredicate<JavaClass>("sealed wire interfaces") {
              @Override
              public boolean test(JavaClass input) {
                return input.isInterface()
                    && input.isSealed()
                    && input.isAnnotatedWith("dev.erst.gridgrind.core.metadata.Union");
              }
            })
        .should(USE_RECORD_WIRE_VARIANTS)
        .because("each concrete wire alternative is one record contract");
  }

  private static Stream<JavaClass> wireVariants(JavaClass union) {
    return union.getPermittedSubclasses().orElseThrow().stream()
        .flatMap(
            variant ->
                variant.isInterface() && variant.isSealed()
                    ? wireVariants(variant)
                    : Stream.of(variant));
  }

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
        .and(wireUnionsUseRecordVariants())
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
