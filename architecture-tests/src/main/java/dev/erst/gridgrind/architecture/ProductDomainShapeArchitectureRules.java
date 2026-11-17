package dev.erst.gridgrind.architecture;

import static com.tngtech.archunit.lang.syntax.ArchRuleDefinition.classes;

import com.tngtech.archunit.base.DescribedPredicate;
import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.lang.ArchCondition;
import com.tngtech.archunit.lang.ArchRule;
import com.tngtech.archunit.lang.ConditionEvents;
import com.tngtech.archunit.lang.SimpleConditionEvent;
import java.util.stream.Stream;

/** Protects the closed wire shapes consumed by record metadata and JSON codecs. */
enum ProductDomainShapeArchitectureRules {
  ;

  static ArchRule unionTagsUseTypeOrStatus() {
    return classes()
        .that()
        .areAnnotatedWith(dev.erst.gridgrind.core.metadata.Union.class)
        .should(
            new ArchCondition<JavaClass>("use type or status as the union discriminator") {
              @Override
              public void check(JavaClass item, ConditionEvents events) {
                String discriminator =
                    item.getAnnotationOfType(dev.erst.gridgrind.core.metadata.Union.class)
                        .discriminator();
                boolean structuralStep =
                    discriminator.isEmpty()
                        && "dev.erst.gridgrind.core.plan.WorkbookStep".equals(item.getName());
                if (!structuralStep
                    && !java.util.Set.of("type", "status").contains(discriminator)) {
                  events.add(
                      SimpleConditionEvent.violated(
                          item,
                          item.getName()
                              + " declares unsupported discriminator: "
                              + discriminator));
                }
              }
            })
        .because("wire identity uses type and completed-process outcomes use status");
  }

  private static final ArchCondition<JavaClass> USE_RECORD_WIRE_VARIANTS =
      new ArchCondition<>("use records for every concrete wire variant") {
        @Override
        public void check(JavaClass item, ConditionEvents events) {
          wireVariants(item)
              .filter(variant -> !variant.isRecord())
              .forEach(variant ->
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
        .flatMap(variant ->
            variant.isInterface() && variant.isSealed()
                ? wireVariants(variant)
                : Stream.of(variant));
  }
}
