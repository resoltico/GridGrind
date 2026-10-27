package dev.erst.gridgrind.engine.poi;

import java.lang.invoke.MethodType;
import java.util.Objects;
import org.apache.poi.Version;

/** Describes one private Apache POI contract that GridGrind currently depends on. */
@SuppressWarnings("PMD.CommentRequired")
public record PrivateContract(
    Class<?> owner, String lookupName, String memberSignature, String affectedSurface) {
  public PrivateContract {
    Objects.requireNonNull(owner, "owner must not be null");
    Objects.requireNonNull(lookupName, "lookupName must not be null");
    Objects.requireNonNull(memberSignature, "memberSignature must not be null");
    Objects.requireNonNull(affectedSurface, "affectedSurface must not be null");
  }

  public static PrivateContract field(Class<?> owner, String fieldName, String affectedSurface) {
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    return new PrivateContract(
        owner, fieldName, owner.getName() + "." + fieldName, affectedSurface);
  }

  public static PrivateContract virtualMethod(
      Class<?> owner, String methodName, MethodType methodType, String affectedSurface) {
    Objects.requireNonNull(methodName, "methodName must not be null");
    Objects.requireNonNull(methodType, "methodType must not be null");
    return new PrivateContract(
        owner,
        methodName,
        owner.getName() + "#" + methodName + parameterSignature(methodType),
        affectedSurface);
  }

  public static PrivateContract constructor(
      Class<?> owner, MethodType constructorType, String affectedSurface) {
    Objects.requireNonNull(constructorType, "constructorType must not be null");
    return new PrivateContract(
        owner, "<init>", owner.getName() + parameterSignature(constructorType), affectedSurface);
  }

  public String failureMessage() {
    return "Apache POI private contract unavailable: "
        + memberSignature
        + " required for "
        + affectedSurface
        + ". GridGrind currently targets Apache POI "
        + Version.getVersion()
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
}
