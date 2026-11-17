package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolContractSupport;
import dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestExplicitNullField;
import dev.erst.gridgrind.core.problem.RequestMalformedScalar;
import dev.erst.gridgrind.core.problem.RequestMissingTypeDiscriminator;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import dev.erst.gridgrind.core.problem.RequestUnknownTypeDiscriminator;
import dev.erst.gridgrind.core.problem.RequestUnsupportedEnumValue;
import dev.erst.gridgrind.core.vocabulary.WireScalar;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

/** Checks original tolerant nodes against the model's field and union declarations. */
final class RequestSchemaStructure {
  private RequestSchemaStructure() {}

  static void validate(RequestJsonNode root, List<RequestStructuralProblem> problems) {
    validate(root, WorkbookPlan.class, "", root.byteOffset(), problems);
  }

  static void validateRecord(
      RequestJsonNode node,
      Class<? extends Record> type,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    object(node, type, path, offset, true, problems);
  }

  private static void validate(
      RequestJsonNode node,
      Type type,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (node instanceof RequestJsonNull) {
      problems.add(new RequestExplicitNullField(location(path), offset));
      return;
    }
    Type unwrapped = unwrapped(type);
    Class<?> raw = RequestJavaTypes.rawType(unwrapped);
    if (WireScalar.isValueType(raw)) raw = WireScalar.wireType(raw);
    if (raw == List.class) {
      array(node, RequestJavaTypes.typeArgument(unwrapped), path, offset, problems);
    } else if (raw.isSealed()) {
      union(node, raw, path, offset, problems);
    } else if (raw.isRecord()) {
      object(node, raw.asSubclass(Record.class), path, offset, true, problems);
    } else if (raw == String.class || raw.isEnum() || RequestJavaTypes.isTemporalType(raw)) {
      string(node, raw, path, offset, problems);
    } else if (raw == boolean.class || raw == Boolean.class) {
      if (!(node instanceof RequestJsonBoolean)) {
        problems.add(new RequestMalformedScalar(location(path), "a JSON boolean", offset));
      }
    } else if (RequestJavaTypes.isNumberType(raw)) {
      RequestScalarTokens.validate(node, raw, location(path), offset, problems);
    } else {
      throw new IllegalStateException("Unsupported model structure at " + path + ": " + type);
    }
  }

  private static void object(
      RequestJsonNode node,
      Class<? extends Record> type,
      String path,
      long offset,
      boolean unknowns,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonObject object)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON object", offset));
      return;
    }
    var members = RequestObjectMembers.index(object);
    RequestObjectMembers.collectMissingRequiredFields(members, path, required(type), problems);
    if (unknowns) RequestObjectMembers.collectUnknownFields(members, path, fields(type), problems);
    for (RecordComponent component : components(type)) {
      for (RequestJsonMember member : members.membersNamed(component.getName())) {
        validate(
            member.value(),
            component.getGenericType(),
            RequestJsonPaths.field(path, component.getName()),
            member.nameByteOffset(),
            problems);
      }
    }
  }

  private static void array(
      RequestJsonNode node,
      Type item,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonArray array)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON array", offset));
      return;
    }
    for (int index = 0; index < array.elements().size(); index++) {
      RequestJsonNode element = array.elements().get(index);
      validate(element, item, path + "[" + index + "]", element.byteOffset(), problems);
    }
  }

  private static void string(
      RequestJsonNode node,
      Class<?> raw,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonString text)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON string", offset));
      return;
    }
    if (raw.isEnum()) {
      List<String> choices =
          Arrays.stream(raw.getEnumConstants()).map(value -> ((Enum<?>) value).name()).toList();
      if (!choices.contains(text.value())) {
        problems.add(
            new RequestUnsupportedEnumValue(location(path), text.value(), choices, offset));
      }
    } else if (RequestJavaTypes.isTemporalType(raw)) {
      RequestScalarTokens.validate(node, raw, location(path), offset, problems);
    }
  }

  private static void union(
      RequestJsonNode node,
      Class<?> root,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonObject object)) {
      problems.add(new RequestMalformedScalar(location(path), "a JSON object", offset));
      return;
    }
    List<Class<? extends Record>> variants =
        ProtocolTypeMetadataSupport.declaredVariants(root).stream()
            .sorted(Comparator.comparing(ProtocolTypeMetadataSupport::requiredTypeId))
            .toList();
    String discriminator = ProtocolTypeMetadataSupport.discriminatorField(root);
    if (discriminator.isEmpty()) {
      structuralUnion(object, root, variants, path, offset, problems);
      return;
    }
    String fieldPath = RequestJsonPaths.field(path, discriminator);
    var members = RequestObjectMembers.index(object);
    List<RequestJsonMember> tags = members.membersNamed(discriminator);
    Set<Class<? extends Record>> viable = new LinkedHashSet<>();
    if (tags.isEmpty()) problems.add(new RequestMissingTypeDiscriminator(fieldPath));
    for (RequestJsonMember tag : tags) {
      if (tag.value() instanceof RequestJsonNull) {
        problems.add(new RequestExplicitNullField(fieldPath, tag.nameByteOffset()));
      } else if (!(tag.value() instanceof RequestJsonString text)) {
        problems.add(
            new RequestMalformedScalar(fieldPath, "a JSON string type id", tag.nameByteOffset()));
      } else {
        var selected =
            variants.stream()
                .filter(variant ->
                    text.value().equals(ProtocolTypeMetadataSupport.requiredTypeId(variant)))
                .findFirst();
        if (selected.isPresent()) {
          viable.add(selected.orElseThrow());
        } else {
          Optional<Class<?>> similarity =
              root == Selector.class ? Optional.empty() : Optional.of(root);
          problems.add(
              new RequestUnknownTypeDiscriminator(
                  fieldPath,
                  text.value(),
                  similarity
                      .map(type ->
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
      variants.forEach(variant -> allowed.addAll(fields(variant)));
    } else {
      allowed.addAll(fields(viable.iterator().next()));
      viable.forEach(variant -> allowed.retainAll(fields(variant)));
    }
    RequestObjectMembers.collectUnknownFields(members, path, allowed, problems);
    List<RequestStructuralProblem> found = new ArrayList<>();
    for (Class<? extends Record> variant : viable) {
      found.clear();
      object(object, variant, path, offset, false, found);
      found.stream().filter(problem -> !problems.contains(problem)).forEach(problems::add);
    }
  }

  private static void structuralUnion(
      RequestJsonObject object,
      Class<?> root,
      List<Class<? extends Record>> variants,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    Set<String> common = new LinkedHashSet<>(required(variants.getFirst()));
    variants.forEach(variant -> common.retainAll(required(variant)));
    Set<String> all = new LinkedHashSet<>();
    variants.forEach(variant -> all.addAll(fields(variant)));
    var members = RequestObjectMembers.index(object);
    RequestObjectMembers.collectMissingRequiredFields(members, path, List.copyOf(common), problems);
    RequestObjectMembers.collectUnknownFields(members, path, all, problems);
    List<String> payloads = new ArrayList<>();
    for (int index = 0; index < variants.size(); index++) {
      Class<? extends Record> variant = variants.get(index);
      for (String field : required(variant)) if (!common.contains(field)) payloads.add(field);
      for (RecordComponent component : components(variant)) {
        if (index != 0 && common.contains(component.getName())) continue;
        for (RequestJsonMember member : members.membersNamed(component.getName())) {
          validate(
              member.value(),
              component.getGenericType(),
              RequestJsonPaths.field(path, component.getName()),
              member.nameByteOffset(),
              problems);
        }
      }
    }
    if (payloads.stream().filter(field -> members.member(field).isPresent()).count() != 1) {
      List<String> ordered =
          ProtocolTypeMetadataSupport.declaredVariants(root).stream()
              .flatMap(type -> required(type).stream())
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

  private static List<RecordComponent> components(Class<? extends Record> type) {
    return GridGrindProtocolContractSupport.visibleRecordComponents(type);
  }

  private static List<String> required(Class<? extends Record> type) {
    List<String> names = new ArrayList<>(GridGrindProtocolContractSupport.requiredFieldNames(type));
    String discriminator = ProtocolTypeMetadataSupport.discriminatorField(type);
    if (!discriminator.isEmpty()) names.add(discriminator);
    return List.copyOf(names);
  }

  private static Set<String> fields(Class<? extends Record> type) {
    Set<String> names = new LinkedHashSet<>();
    components(type).forEach(component -> names.add(component.getName()));
    String discriminator = ProtocolTypeMetadataSupport.discriminatorField(type);
    if (!discriminator.isEmpty()) names.add(discriminator);
    return names;
  }

  private static Type unwrapped(Type type) {
    return RequestJavaTypes.rawType(type) == Optional.class
        ? RequestJavaTypes.typeArgument(type)
        : type;
  }

  private static String location(String path) {
    return path.isEmpty() ? "request" : path;
  }
}
