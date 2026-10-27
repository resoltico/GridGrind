package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.metadata.ModelGraph;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestExplicitNullField;
import dev.erst.gridgrind.core.problem.RequestMalformedScalar;
import dev.erst.gridgrind.core.problem.RequestMissingTypeDiscriminator;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import dev.erst.gridgrind.core.problem.RequestUnknownTypeDiscriminator;
import dev.erst.gridgrind.core.problem.RequestUnsupportedEnumValue;
import dev.erst.gridgrind.core.vocabulary.WireScalar;
import dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport;
import dev.erst.gridgrind.json.schema.JsonSchemas;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;
import tools.jackson.databind.JsonNode;

/** Checks emitted schema structures on the original tree without collapsing duplicate keys. */
final class RequestSchemaStructure {
  private static final JsonNode DOCUMENT =
      GridGrindJsonMapperSupport.JSON_MAPPER.readTree(
          JsonSchemas.document(JsonSchemas.Document.REQUEST));
  private static final Map<String, Class<?>> TYPES =
      ModelGraph.reachableTypes(WorkbookPlan.class).stream()
          .collect(Collectors.toUnmodifiableMap(Class::getName, type -> type));

  private RequestSchemaStructure() {}

  static void validate(RequestJsonNode root, List<RequestStructuralProblem> problems) {
    validate(root, DOCUMENT, WorkbookPlan.class, "", root.byteOffset(), problems);
  }

  static void validateRecord(
      RequestJsonNode node,
      Class<? extends Record> type,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    object(node, DOCUMENT.path("$defs").path(type.getName()), type, path, offset, true, problems);
  }

  private static void validate(
      RequestJsonNode node,
      JsonNode schema,
      Type javaType,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (node instanceof RequestJsonNull) {
      problems.add(new RequestExplicitNullField(location(path), offset));
      return;
    }
    if (schema.has("$ref")) {
      String name = schema.path("$ref").asString().substring("#/$defs/".length());
      Class<?> type = TYPES.getOrDefault(name, Object.class);
      if (type.isRecord()) {
        validateRecord(node, type.asSubclass(Record.class), path, offset, problems);
      } else {
        validate(node, DOCUMENT.path("$defs").path(name), type, path, offset, problems);
      }
      return;
    }
    if (schema.has("oneOf")) {
      union(node, schema.path("oneOf"), javaType, path, offset, problems);
      return;
    }
    switch (schema.path("type").asString()) {
      case "object" -> object(node, schema, javaType, path, offset, true, problems);
      case "array" -> array(node, schema, javaType, path, offset, problems);
      case "string" -> string(node, schema, javaType, path, offset, problems);
      case "boolean" -> {
        if (!(node instanceof RequestJsonBoolean)) {
          problems.add(new RequestMalformedScalar(location(path), "a JSON boolean", offset));
        }
      }
      case "number", "integer" -> scalarNumber(node, javaType, path, offset, problems);
      default -> throw new IllegalStateException("Unsupported emitted schema structure at " + path);
    }
  }

  private static void object(
      RequestJsonNode node,
      JsonNode schema,
      Type javaType,
      String path,
      long offset,
      boolean unknowns,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonObject object)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON object", offset));
      return;
    }
    var members = RequestObjectMembers.index(object);
    RequestObjectMembers.collectMissingRequiredFields(
        members, path, strings(schema.path("required")), problems);
    JsonNode properties = schema.path("properties");
    if (unknowns) {
      RequestObjectMembers.collectUnknownFields(members, path, fields(properties), problems);
    }
    for (var entry : properties.properties()) {
      if (entry.getValue().has("const")) {
        continue;
      }
      Type component = component(javaType, entry.getKey());
      for (RequestJsonMember member : members.membersNamed(entry.getKey())) {
        validate(
            member.value(),
            entry.getValue(),
            component,
            RequestJsonPaths.field(path, entry.getKey()),
            member.nameByteOffset(),
            problems);
      }
    }
  }

  private static void array(
      RequestJsonNode node,
      JsonNode schema,
      Type javaType,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonArray array)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON array", offset));
      return;
    }
    Type item = RequestJavaTypes.typeArgument(unwrapped(javaType));
    for (int index = 0; index < array.elements().size(); index++) {
      RequestJsonNode element = array.elements().get(index);
      validate(
          element,
          schema.path("items"),
          item,
          path + "[" + index + "]",
          element.byteOffset(),
          problems);
    }
  }

  private static void string(
      RequestJsonNode node,
      JsonNode schema,
      Type javaType,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonString text)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON string", offset));
      return;
    }
    Class<?> raw = RequestJavaTypes.rawType(unwrapped(javaType));
    if (raw.isEnum()
        && schema.has("enum")
        && !strings(schema.path("enum")).contains(text.value())) {
      problems.add(
          new RequestUnsupportedEnumValue(
              location(path), text.value(), strings(schema.path("enum")), offset));
      return;
    }
    if (RequestJavaTypes.isTemporalType(raw)) {
      RequestScalarTokens.validate(node, raw, location(path), offset, problems);
    }
  }

  private static void scalarNumber(
      RequestJsonNode node,
      Type javaType,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    Class<?> raw = RequestJavaTypes.rawType(unwrapped(javaType));
    if (WireScalar.isValueType(raw)) {
      raw = raw.getRecordComponents()[0].getType();
    }
    RequestScalarTokens.validate(node, raw, location(path), offset, problems);
  }

  private static void union(
      RequestJsonNode node,
      JsonNode alternatives,
      Type javaType,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonObject object)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON object", offset));
      return;
    }
    List<JsonNode> variants = new ArrayList<>();
    for (JsonNode alternative : alternatives) {
      variants.add(dereference(alternative));
    }
    Optional<String> discriminator =
        variants.getFirst().path("properties").properties().stream()
            .filter(entry -> entry.getValue().has("const"))
            .map(Map.Entry::getKey)
            .filter(
                name ->
                    variants.stream()
                        .allMatch(variant -> variant.path("properties").path(name).has("const")))
            .findFirst();
    if (discriminator.isEmpty()) {
      structuralUnion(object, alternatives, variants, javaType, path, offset, problems);
      return;
    }
    String field = discriminator.orElseThrow();
    String fieldPath = RequestJsonPaths.field(path, field);
    var members = RequestObjectMembers.index(object);
    List<RequestJsonMember> tags = members.membersNamed(field);
    Set<JsonNode> viable = new LinkedHashSet<>();
    if (tags.isEmpty()) {
      problems.add(new RequestMissingTypeDiscriminator(fieldPath));
    }
    for (RequestJsonMember tag : tags) {
      if (tag.value() instanceof RequestJsonNull) {
        problems.add(new RequestExplicitNullField(fieldPath, tag.nameByteOffset()));
      } else if (!(tag.value() instanceof RequestJsonString text)) {
        problems.add(
            new RequestMalformedScalar(fieldPath, "a JSON string type id", tag.nameByteOffset()));
      } else {
        var selected =
            variants.stream()
                .filter(
                    variant ->
                        text.value()
                            .equals(
                                variant.path("properties").path(field).path("const").asString()))
                .findFirst();
        if (selected.isPresent()) {
          viable.add(selected.orElseThrow());
        } else {
          Class<?> root = RequestJavaTypes.rawType(javaType);
          Optional<Class<?>> similarity =
              root == Selector.class ? Optional.empty() : Optional.of(root);
          problems.add(
              new RequestUnknownTypeDiscriminator(
                  fieldPath,
                  text.value(),
                  similarity
                      .map(
                          type ->
                              GridGrindJsonSubtypeProblemSupport.similarTypeIds(type, text.value()))
                      .orElse(List.of()),
                  GridGrindJsonSubtypeProblemSupport.specificGuidance(
                      fieldPath, text.value(), similarity),
                  tag.nameByteOffset()));
        }
      }
    }
    Set<String> allowed = new LinkedHashSet<>();
    if (viable.isEmpty()) {
      variants.forEach(variant -> allowed.addAll(fields(variant.path("properties"))));
    } else {
      allowed.addAll(fields(viable.iterator().next().path("properties")));
      viable.forEach(variant -> allowed.retainAll(fields(variant.path("properties"))));
    }
    RequestObjectMembers.collectUnknownFields(members, path, allowed, problems);
    List<RequestStructuralProblem> found = new ArrayList<>();
    for (JsonNode variant : viable) {
      found.clear();
      int index = variants.indexOf(variant);
      Type concrete = referencedType(alternatives.get(index));
      object(object, variant, concrete, path, offset, false, found);
      found.stream().filter(problem -> !problems.contains(problem)).forEach(problems::add);
    }
  }

  private static void structuralUnion(
      RequestJsonObject object,
      JsonNode alternatives,
      List<JsonNode> variants,
      Type javaType,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    Set<String> common = new LinkedHashSet<>(strings(variants.getFirst().path("required")));
    variants.forEach(variant -> common.retainAll(strings(variant.path("required"))));
    Set<String> all = new LinkedHashSet<>();
    variants.forEach(variant -> all.addAll(fields(variant.path("properties"))));
    var members = RequestObjectMembers.index(object);
    RequestObjectMembers.collectMissingRequiredFields(members, path, List.copyOf(common), problems);
    RequestObjectMembers.collectUnknownFields(members, path, all, problems);
    List<String> payloads = new ArrayList<>();
    for (int index = 0; index < variants.size(); index++) {
      JsonNode variant = variants.get(index);
      Type concrete = referencedType(alternatives.get(index));
      for (String field : strings(variant.path("required"))) {
        if (!common.contains(field)) {
          payloads.add(field);
        }
      }
      for (var entry : variant.path("properties").properties()) {
        if (index != 0 && common.contains(entry.getKey())) {
          continue;
        }
        for (var member : members.membersNamed(entry.getKey())) {
          validate(
              member.value(),
              entry.getValue(),
              component(concrete, entry.getKey()),
              RequestJsonPaths.field(path, entry.getKey()),
              member.nameByteOffset(),
              problems);
        }
      }
    }
    if (payloads.stream().filter(field -> members.member(field).isPresent()).count() != 1) {
      List<String> ordered =
          dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport.declaredVariants(
                  RequestJavaTypes.rawType(javaType))
              .stream()
              .flatMap(
                  type ->
                      strings(DOCUMENT.path("$defs").path(type.getName()).path("required"))
                          .stream())
              .filter(payloads::contains)
              .distinct()
              .toList();
      problems.add(
          new RequestMalformedScalar(
              location(path),
              "an object containing exactly one of "
                  + String.join(", ", ordered.subList(0, ordered.size() - 1))
                  + ", or "
                  + ordered.getLast(),
              offset));
    }
  }

  private static JsonNode dereference(JsonNode schema) {
    return schema.has("$ref")
        ? DOCUMENT.path("$defs").path(schema.path("$ref").asString().substring("#/$defs/".length()))
        : schema;
  }

  private static Type referencedType(JsonNode schema) {
    return TYPES.getOrDefault(
        schema.path("$ref").asString().substring("#/$defs/".length()), Object.class);
  }

  private static Type component(Type owner, String name) {
    Class<?> raw = RequestJavaTypes.rawType(unwrapped(owner));
    if (!raw.isRecord()) {
      throw new IllegalStateException("Schema object lacks its canonical record: " + owner);
    }
    for (RecordComponent component : raw.getRecordComponents()) {
      if (component.getName().equals(name)) {
        return component.getGenericType();
      }
    }
    throw new IllegalStateException("Schema field lacks its canonical component: " + name);
  }

  private static Type unwrapped(Type type) {
    return RequestJavaTypes.rawType(type) == Optional.class
        ? RequestJavaTypes.typeArgument(type)
        : type;
  }

  private static List<String> strings(JsonNode array) {
    List<String> result = new ArrayList<>();
    array.forEach(value -> result.add(value.asString()));
    return List.copyOf(result);
  }

  private static Set<String> fields(JsonNode properties) {
    return properties.properties().stream()
        .map(Map.Entry::getKey)
        .collect(Collectors.toUnmodifiableSet());
  }

  private static String location(String path) {
    return path.isEmpty() ? "request" : path;
  }
}
