package dev.erst.gridgrind.excel;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.util.Objects;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;

/** Package-owned seam for POI relation removal, including reflective lookup hardening. */
final class PoiRelationRemoval {
  private PoiRelationRemoval() {}

  static BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> defaultRemover() {
    MethodHandle methodHandle = removePoiRelationMethod(MethodHandles.lookup());
    return (parent, child) -> invokePoiRelationRemoval(methodHandle, parent, child);
  }

  static MethodHandle removePoiRelationMethod(MethodHandles.Lookup lookup) {
    Objects.requireNonNull(lookup, "lookup must not be null");
    try {
      return MethodHandles.privateLookupIn(POIXMLDocumentPart.class, lookup)
          .findVirtual(
              POIXMLDocumentPart.class,
              "removeRelation",
              MethodType.methodType(boolean.class, POIXMLDocumentPart.class, boolean.class));
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException("Failed to resolve POI removeRelation handle", exception);
    }
  }

  static boolean invokePoiRelationRemoval(
      MethodHandle methodHandle, POIXMLDocumentPart parent, POIXMLDocumentPart child) {
    Objects.requireNonNull(methodHandle, "methodHandle must not be null");
    Objects.requireNonNull(parent, "parent must not be null");
    Objects.requireNonNull(child, "child must not be null");
    try {
      return (boolean) methodHandle.invoke(parent, child, true);
    } catch (Throwable exception) {
      throw new IllegalStateException("Failed to remove POI relation", exception);
    }
  }
}
