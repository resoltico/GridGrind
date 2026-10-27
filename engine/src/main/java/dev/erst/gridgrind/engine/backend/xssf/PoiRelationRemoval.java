package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.engine.poi.PoiAccess;
import dev.erst.gridgrind.engine.poi.PrivateContract;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.util.Objects;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;

/** Package-owned seam for POI relation removal, including reflective lookup hardening. */
@SuppressWarnings("PMD.CommentRequired")
public final class PoiRelationRemoval {
  static final PrivateContract REMOVE_RELATION_CONTRACT =
      PrivateContract.virtualMethod(
          POIXMLDocumentPart.class,
          "removeRelation",
          MethodType.methodType(boolean.class, POIXMLDocumentPart.class, boolean.class),
          "drawing, chart, and pivot relation cleanup");

  private PoiRelationRemoval() {}

  public static BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> defaultRemover() {
    BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> invoker =
        removePoiRelationInvoker(MethodHandles.lookup());
    return (parent, child) -> invokePoiRelationRemoval(invoker, parent, child);
  }

  static BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> removePoiRelationInvoker(
      MethodHandles.Lookup lookup) {
    MethodHandle methodHandle =
        PoiAccess.requireVirtual(
            lookup,
            REMOVE_RELATION_CONTRACT,
            MethodType.methodType(boolean.class, POIXMLDocumentPart.class, boolean.class));
    MethodHandle typedInvoker =
        MethodHandles.insertArguments(methodHandle, 2, true)
            .asType(
                MethodType.methodType(
                    boolean.class, POIXMLDocumentPart.class, POIXMLDocumentPart.class));
    return relationInvoker(typedInvoker);
  }

  static boolean invokePoiRelationRemoval(
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> invoker,
      POIXMLDocumentPart parent,
      POIXMLDocumentPart child) {
    Objects.requireNonNull(invoker, "invoker must not be null");
    Objects.requireNonNull(parent, "parent must not be null");
    Objects.requireNonNull(child, "child must not be null");
    try {
      return invoker.test(parent, child);
    } catch (RuntimeException exception) {
      throw new IllegalStateException("Failed to remove POI relation", exception);
    }
  }

  @SuppressWarnings("unchecked")
  private static BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> relationInvoker(
      MethodHandle typedInvoker) {
    return PoiAccess.asInterfaceInstance(BiPredicate.class, typedInvoker);
  }
}
