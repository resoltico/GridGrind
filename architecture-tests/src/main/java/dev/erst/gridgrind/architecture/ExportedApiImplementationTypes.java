package dev.erst.gridgrind.architecture;

import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.domain.JavaMember;
import com.tngtech.archunit.core.domain.JavaModifier;
import java.util.Set;
import java.util.TreeSet;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/** Extracts implementation types that are externally reachable through one exported API class. */
final class ExportedApiImplementationTypes {
  Set<String> namesFor(JavaClass apiClass) {
    EngineImplementationTypeClassifier classifier = new EngineImplementationTypeClassifier();
    Stream<JavaClass> memberTypes =
        apiClass.getAllMembers().stream()
            .filter(member -> isExternallyReachable(apiClass, member))
            .flatMap(member -> member.getAllInvolvedRawTypes().stream());
    return Stream.concat(
            Stream.concat(
                apiClass.getAllRawSuperclasses().stream(), apiClass.getAllRawInterfaces().stream()),
            memberTypes)
        .filter(classifier::isImplementationType)
        .map(JavaClass::getName)
        .collect(Collectors.toCollection(TreeSet::new));
  }

  boolean isExternallyReachable(JavaClass owner, JavaMember member) {
    return member.getModifiers().contains(JavaModifier.PUBLIC)
        || (!owner.getModifiers().contains(JavaModifier.FINAL)
            && member.getModifiers().contains(JavaModifier.PROTECTED));
  }
}
