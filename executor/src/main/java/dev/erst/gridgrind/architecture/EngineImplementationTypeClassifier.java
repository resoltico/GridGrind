package dev.erst.gridgrind.architecture;

import com.tngtech.archunit.core.domain.JavaClass;

/** Classifies engine-internal types that must not cross the exported API boundary. */
final class EngineImplementationTypeClassifier {
  boolean isImplementationType(JavaClass javaClass) {
    String packageName = javaClass.getPackageName();
    return isInOrBelow(packageName, "dev.erst.gridgrind.engine.runtime")
        || (isInOrBelow(packageName, "dev.erst.gridgrind.excel")
            && !isInOrBelow(packageName, "dev.erst.gridgrind.excel.foundation"));
  }

  static boolean isInOrBelow(String packageName, String parentPackage) {
    return packageName.equals(parentPackage) || packageName.startsWith(parentPackage + ".");
  }
}
