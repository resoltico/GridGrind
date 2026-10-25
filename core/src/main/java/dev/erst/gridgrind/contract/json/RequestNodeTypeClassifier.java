package dev.erst.gridgrind.contract.json;

import java.lang.reflect.Type;
import java.util.Collection;
import java.util.Optional;

/** Identifies container creator types whose elements require recursive request validation. */
final class RequestNodeTypeClassifier {
  private RequestNodeTypeClassifier() {}

  static boolean isOptional(Type type) {
    return RequestTypeSupport.rawType(type) == Optional.class;
  }

  static boolean isCollection(Type type) {
    return Collection.class.isAssignableFrom(RequestTypeSupport.rawType(type));
  }
}
