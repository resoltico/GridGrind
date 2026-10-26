package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.core.metadata.ModelDefaults;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.formula.FormulaInputException;
import dev.erst.gridgrind.core.problem.FormulaRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestProblemDescriptor;
import dev.erst.gridgrind.core.problem.RequestProblemSource;
import dev.erst.gridgrind.core.vocabulary.WireScalar;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/** Exercises canonical model invariants on shape-valid input before the wire binder runs. */
final class RequestConstructorValidation {
  private RequestConstructorValidation() {}

  static <T> Optional<T> validate(
      RequestJsonNode node, Class<T> type, String path, List<RequestBindingFailure> problems) {
    return value(node, type, path, problems).map(type::cast);
  }

  private static Optional<Object> value(
      RequestJsonNode node, Type type, String path, List<RequestBindingFailure> problems) {
    if (type instanceof ParameterizedType parameterized) {
      Type item = parameterized.getActualTypeArguments()[0];
      if (parameterized.getRawType() == Optional.class) {
        return value(node, item, path, problems).map(Optional::of);
      }
      if (parameterized.getRawType() == List.class && node instanceof RequestJsonArray array) {
        List<Object> items = new ArrayList<>();
        int count = problems.size();
        for (int index = 0; index < array.elements().size(); index++) {
          value(array.elements().get(index), item, path + "[" + index + "]", problems)
              .ifPresent(items::add);
        }
        return problems.size() == count ? Optional.of(List.copyOf(items)) : Optional.empty();
      }
      throw new IllegalStateException("Unsupported model container: " + type);
    }
    Class<?> raw = RequestJavaTypes.rawType(type);
    if (WireScalar.isValueType(raw)) {
      RecordComponent component = raw.getRecordComponents()[0];
      Object scalar = scalar(node, component.getType());
      return construct(raw, new Object[] {scalar}, path, true, problems);
    }
    if (raw.isSealed()) {
      return value(node, variant((RequestJsonObject) node, raw), path, problems);
    }
    if (raw.isRecord()) {
      return record((RequestJsonObject) node, raw, path, problems);
    }
    return Optional.of(scalar(node, raw));
  }

  private static Class<?> variant(RequestJsonObject node, Class<?> root) {
    List<Class<? extends Record>> variants = ProtocolTypeMetadataSupport.declaredVariants(root);
    Optional<String> discriminator = discriminator(root);
    if (discriminator.isPresent()) {
      String name = discriminator.orElseThrow();
      RequestJsonString tag =
          (RequestJsonString) RequestObjectMembers.index(node).member(name).orElseThrow().value();
      return ProtocolTypeMetadataSupport.subtypeForTypeId(root, tag.value()).orElseThrow();
    }
    var members = RequestObjectMembers.index(node);
    return variants.stream()
        .filter(
            type ->
                java.util.Arrays.stream(type.getRecordComponents())
                    .allMatch(
                        component ->
                            members.member(component.getName()).isPresent()
                                || component.getType() == Optional.class
                                || component.isAnnotationPresent(WireDefault.class)))
        .findFirst()
        .orElseThrow(
            () -> new IllegalStateException("Shape-valid union has no canonical alternative"));
  }

  private static Optional<String> discriminator(Class<?> root) {
    Union annotation = root.getAnnotation(Union.class);
    if (annotation != null) {
      return annotation.discriminator().isEmpty()
          ? Optional.empty()
          : Optional.of(annotation.discriminator());
    }
    if (!root.isSealed()) {
      return Optional.empty();
    }
    for (Class<?> permitted : root.getPermittedSubclasses()) {
      Optional<String> inherited = discriminator(permitted);
      if (inherited.isPresent()) {
        return inherited;
      }
    }
    return Optional.empty();
  }

  @SuppressWarnings(
      "PMD.UseConcurrentHashMap") // Each record's sibling values are local to its construction.
  private static Optional<Object> record(
      RequestJsonObject node, Class<?> type, String path, List<RequestBindingFailure> problems) {
    var members = RequestObjectMembers.index(node);
    RecordComponent[] components = type.getRecordComponents();
    Map<String, Object> values = new LinkedHashMap<>();
    int before = problems.size();
    for (RecordComponent component : components) {
      var member = members.member(component.getName());
      if (member.isPresent()) {
        value(
                member.orElseThrow().value(),
                component.getGenericType(),
                RequestJsonPaths.field(path, component.getName()),
                problems)
            .ifPresent(decoded -> values.put(component.getName(), decoded));
      } else if (component.getType() == Optional.class) {
        values.put(component.getName(), Optional.empty());
      }
    }
    if (problems.size() != before) {
      return Optional.empty();
    }
    while (values.size() < components.length) {
      int count = values.size();
      for (RecordComponent component : components) {
        if (values.containsKey(component.getName())) {
          continue;
        }
        WireDefault omission = component.getAnnotation(WireDefault.class);
        if (omission == null) {
          throw new IllegalStateException(
              "Shape-valid record lacks required component: " + component);
        }
        if (omission.from().isEmpty() || values.containsKey(omission.from())) {
          values.put(component.getName(), ModelDefaults.value(component, values));
        }
      }
      if (values.size() == count) {
        throw new IllegalStateException("Model defaults contain a cycle: " + type);
      }
    }
    Object[] arguments =
        java.util.Arrays.stream(components)
            .map(component -> values.get(component.getName()))
            .toArray();
    return construct(type, arguments, path, false, problems);
  }

  private static Optional<Object> construct(
      Class<?> type,
      Object[] arguments,
      String path,
      boolean scalar,
      List<RequestBindingFailure> problems) {
    try {
      Class<?>[] parameters =
          java.util.Arrays.stream(type.getRecordComponents())
              .map(RecordComponent::getType)
              .toArray(Class<?>[]::new);
      return Optional.of(type.getConstructor(parameters).newInstance(arguments));
    } catch (InvocationTargetException wrapped) {
      Throwable cause = wrapped.getCause();
      if (!(cause instanceof IllegalArgumentException invalid)) {
        throw new IllegalStateException(
            "Canonical model validation failed unexpectedly: " + type, cause);
      }
      problems.add(problem(invalid, path, scalar));
      return Optional.empty();
    } catch (ReflectiveOperationException failure) {
      throw new IllegalStateException(
          "Canonical model needs an accessible public constructor: " + type, failure);
    }
  }

  static RequestBindingFailure problem(
      IllegalArgumentException invalid, String path, boolean scalar) {
    String location = path.isEmpty() ? "request" : path;
    if (invalid instanceof FormulaInputException input) {
      if (input.fieldPath().isPresent()) {
        location = relative(path, input.fieldPath().orElseThrow());
      }
      FormulaRequestException formula =
          FormulaRequestProblemSupport.inputFailure(
                  invalid, Optional.of(location), Optional.empty(), Optional.empty(), invalid)
              .orElseThrow();
      return new RequestBindingFailure(formula, location, Optional.empty());
    }
    if (invalid instanceof RequestProblemSource source) {
      Optional<String> own = source.requestProblem().jsonPath();
      if (!scalar && own.isPresent()) {
        location = relative(path, own.orElseThrow());
      }
      IllegalArgumentException normalized =
          switch (source.requestProblem()) {
            case RequestProblemDescriptor.Shape shape ->
                new InvalidRequestShapeException(
                    shape, Optional.of(location), Optional.empty(), Optional.empty(), invalid);
            case RequestProblemDescriptor.Invariant invariant ->
                new InvalidRequestException(
                    invariant, Optional.of(location), Optional.empty(), Optional.empty(), invalid);
          };
      return new RequestBindingFailure(normalized, location, Optional.empty());
    }
    String message =
        Optional.ofNullable(invalid.getMessage())
            .filter(value -> !value.isBlank())
            .orElse("Model value is invalid.");
    return new RequestBindingFailure(
        new InvalidRequestException(
            new MessageInvariant(message, Optional.of(location)),
            Optional.of(location),
            Optional.empty(),
            Optional.empty(),
            invalid),
        location,
        Optional.empty());
  }

  private static String relative(String parent, String field) {
    return parent.isEmpty() ? field : field.startsWith("[") ? parent + field : parent + "." + field;
  }

  @SuppressWarnings({"unchecked", "rawtypes"}) // The declared enum class owns the validated token.
  private static Object scalar(RequestJsonNode node, Class<?> type) {
    if (type == String.class) {
      return ((RequestJsonString) node).value();
    }
    if (type == boolean.class || type == Boolean.class) {
      return ((RequestJsonBoolean) node).value();
    }
    String token =
        node instanceof RequestJsonString text ? text.value() : ((RequestJsonNumber) node).value();
    if (type.isEnum()) return Enum.valueOf((Class<? extends Enum>) type, token);
    if (type == LocalDate.class) return LocalDate.parse(token);
    if (type == LocalDateTime.class) return LocalDateTime.parse(token);
    if (type == byte.class || type == Byte.class) return Byte.valueOf(token);
    if (type == short.class || type == Short.class) return Short.valueOf(token);
    if (type == int.class || type == Integer.class) return Integer.valueOf(token);
    if (type == long.class || type == Long.class) return Long.valueOf(token);
    if (type == float.class || type == Float.class) return Float.valueOf(token);
    if (type == double.class || type == Double.class) return Double.valueOf(token);
    if (type == BigInteger.class) return new BigInteger(token);
    if (type == BigDecimal.class) return new BigDecimal(token);
    throw new IllegalStateException("Unsupported canonical scalar type: " + type);
  }
}
