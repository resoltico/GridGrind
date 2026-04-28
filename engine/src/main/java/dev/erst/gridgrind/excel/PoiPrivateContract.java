package dev.erst.gridgrind.excel;

import java.lang.invoke.MethodType;
import java.lang.module.ModuleDescriptor;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Predicate;

/** Describes one private Apache POI contract that GridGrind currently depends on. */
record PoiPrivateContract(
    Class<?> owner, String lookupName, String memberSignature, String affectedSurface) {
  PoiPrivateContract {
    Objects.requireNonNull(owner, "owner must not be null");
    Objects.requireNonNull(lookupName, "lookupName must not be null");
    Objects.requireNonNull(memberSignature, "memberSignature must not be null");
    Objects.requireNonNull(affectedSurface, "affectedSurface must not be null");
  }

  static PoiPrivateContract field(Class<?> owner, String fieldName, String affectedSurface) {
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    return new PoiPrivateContract(
        owner, fieldName, owner.getName() + "." + fieldName, affectedSurface);
  }

  static PoiPrivateContract virtualMethod(
      Class<?> owner, String methodName, MethodType methodType, String affectedSurface) {
    Objects.requireNonNull(methodName, "methodName must not be null");
    Objects.requireNonNull(methodType, "methodType must not be null");
    return new PoiPrivateContract(
        owner,
        methodName,
        owner.getName() + "#" + methodName + parameterSignature(methodType),
        affectedSurface);
  }

  static PoiPrivateContract constructor(
      Class<?> owner, MethodType constructorType, String affectedSurface) {
    Objects.requireNonNull(constructorType, "constructorType must not be null");
    return new PoiPrivateContract(
        owner, "<init>", owner.getName() + parameterSignature(constructorType), affectedSurface);
  }

  String failureMessage() {
    return "Apache POI private contract unavailable: "
        + memberSignature
        + " required for "
        + affectedSurface
        + ". GridGrind currently targets Apache POI "
        + poiRuntimeVersion(owner)
        + "; update the private POI seam or replace it with a supported POI API.";
  }

  private static String parameterSignature(MethodType methodType) {
    StringBuilder builder = new StringBuilder("(");
    for (int index = 0; index < methodType.parameterCount(); index++) {
      if (index > 0) {
        builder.append(", ");
      }
      builder.append(methodType.parameterType(index).getSimpleName());
    }
    return builder.append(')').toString();
  }

  private static String poiRuntimeVersion(Class<?> owner) {
    String implementationVersion =
        Optional.ofNullable(owner.getPackage())
            .map(Package::getImplementationVersion)
            .filter(Predicate.not(String::isBlank))
            .orElse(null);
    if (implementationVersion != null) {
      return implementationVersion;
    }
    ModuleDescriptor descriptor = owner.getModule().getDescriptor();
    if (descriptor == null) {
      return "unknown";
    }
    return descriptor.rawVersion().orElse("unknown");
  }
}
