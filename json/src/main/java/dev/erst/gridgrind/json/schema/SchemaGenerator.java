package dev.erst.gridgrind.json.schema;

import dev.erst.gridgrind.core.catalog.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.core.catalog.FieldConstraint;
import dev.erst.gridgrind.core.catalog.FieldShape;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolContractSupport;
import dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.core.metadata.CatalogIgnored;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ModelDefaults;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.vocabulary.WireScalar;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.nio.file.Path;
import java.time.temporal.TemporalAccessor;
import java.util.Comparator;
import java.util.Optional;
import java.util.Set;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ObjectNode;

/** Derives document structure and declared scalar constraints from the canonical model. */
final class SchemaGenerator {
  private static final JsonMapper JSON = JsonMapper.builder().build();
  private static final String NON_BLANK_PATTERN = nonBlankPattern();
  private final ObjectNode definitions = JSON.createObjectNode();
  private final boolean observedNumbers;

  private SchemaGenerator(boolean observedNumbers) {
    this.observedNumbers = observedNumbers;
  }

  static ObjectNode generate(Class<?> root, boolean observedNumbers) {
    var generator = new SchemaGenerator(observedNumbers);
    ObjectNode result =
        JSON.createObjectNode()
            .put("$schema", "https://json-schema.org/draft/2020-12/schema")
            .put(
                "$comment",
                "Structure and declared field constraints. GridGrind also enforces value-object, numeric-token, constructor and workflow invariants.");
    result.setAll(generator.node(root));
    result.set("$defs", generator.definitions);
    return result;
  }

  private ObjectNode node(Type type) {
    if (type instanceof ParameterizedType parameterized) {
      Type item = parameterized.getActualTypeArguments()[0];
      if (parameterized.getRawType() == Optional.class) {
        return node(item);
      }
      if (parameterized.getRawType() == java.util.List.class) {
        return JSON.createObjectNode().put("type", "array").set("items", node(item));
      }
      throw new IllegalStateException("Unsupported schema parameterized type: " + type);
    }
    if (!(type instanceof Class<?> raw)) {
      throw new IllegalStateException("Unsupported schema type: " + type);
    }
    if (WireScalar.isValueType(raw)) {
      RecordComponent component = raw.getRecordComponents()[0];
      ObjectNode scalar = node(component.getGenericType());
      metadata(scalar, component, Set.of());
      return scalar;
    }
    if (raw.isEnum()) {
      ObjectNode scalar = JSON.createObjectNode().put("type", "string");
      var candidates = scalar.putArray("enum");
      for (Object value : raw.getEnumConstants()) {
        candidates.add(((Enum<?>) value).name());
      }
      return scalar;
    }
    if (raw == String.class || raw == Path.class || TemporalAccessor.class.isAssignableFrom(raw)) {
      return JSON.createObjectNode().put("type", "string");
    }
    if (raw == boolean.class || raw == Boolean.class) {
      return JSON.createObjectNode().put("type", "boolean");
    }
    if (CatalogFieldMetadataSupport.isNumericType(raw)) {
      return number(raw);
    }
    return reference(raw);
  }

  private ObjectNode number(Class<?> type) {
    boolean integral =
        type == byte.class
            || type == short.class
            || type == int.class
            || type == long.class
            || type == Byte.class
            || type == Short.class
            || type == Integer.class
            || type == Long.class
            || type == java.math.BigInteger.class;
    ObjectNode number = JSON.createObjectNode().put("type", integral ? "integer" : "number");
    if (!observedNumbers
        || !(type == double.class
            || type == Double.class
            || type == float.class
            || type == Float.class)) {
      return number;
    }
    ObjectNode result = JSON.createObjectNode();
    var alternatives = result.putArray("oneOf");
    alternatives.add(number);
    alternatives
        .addObject()
        .put("type", "string")
        .putArray("enum")
        .add("NaN")
        .add("Infinity")
        .add("-Infinity");
    return result;
  }

  private ObjectNode reference(Class<?> type) {
    String id = type.getName();
    if (!definitions.has(id)) {
      ObjectNode definition = definitions.putObject(id);
      if (type.isSealed()) {
        var alternatives = definition.putArray("oneOf");
        ProtocolTypeMetadataSupport.declaredVariants(type).stream()
            .sorted(Comparator.comparing(ProtocolTypeMetadataSupport::requiredTypeId))
            .forEach(variant -> alternatives.add(node(variant)));
      } else if (type.isRecord()) {
        record(definition, type.asSubclass(Record.class));
      } else {
        throw new IllegalStateException("Unsupported schema class: " + type.getName());
      }
    }
    return JSON.createObjectNode().put("$ref", "#/$defs/" + id);
  }

  private void record(ObjectNode schema, Class<? extends Record> type) {
    schema.put("type", "object").put("additionalProperties", false);
    ObjectNode properties = schema.putObject("properties");
    var required = schema.putArray("required");
    Set<String> optional = Set.copyOf(GridGrindProtocolContractSupport.optionalFieldNames(type));
    for (RecordComponent component : type.getRecordComponents()) {
      if (component.isAnnotationPresent(CatalogIgnored.class)
          || component.getAccessor().isAnnotationPresent(CatalogIgnored.class)) {
        continue;
      }
      ObjectNode value = node(component.getGenericType());
      metadata(value, component, optional);
      properties.set(component.getName(), value);
      if (!optional.contains(component.getName())) {
        required.add(component.getName());
      }
    }
    Variant variant = type.getAnnotation(Variant.class);
    if (variant != null) {
      schema.put("description", variant.summary());
      String discriminator = discriminator(type);
      if (!discriminator.isEmpty()) {
        properties.putObject(discriminator).put("const", variant.id());
        required.add(discriminator);
      }
    }
    CatalogType plain = type.getAnnotation(CatalogType.class);
    if (plain != null) {
      schema.put("description", plain.summary());
    }
  }

  private static String discriminator(Class<?> type) {
    for (Class<?> parent : type.getInterfaces()) {
      Union union = parent.getAnnotation(Union.class);
      if (union != null) {
        return union.discriminator();
      }
      String inherited = discriminator(parent);
      if (!inherited.isEmpty()) {
        return inherited;
      }
    }
    return "";
  }

  private void metadata(ObjectNode value, RecordComponent component, Set<String> optional) {
    ModelDefaults.publishedValue(component)
        .ifPresent(defaultValue -> value.set("default", JSON.valueToTree(defaultValue)));
    WireDefault omission = component.getAnnotation(WireDefault.class);
    if (omission != null && !omission.from().isEmpty()) {
      value.put("x-gridgrind-defaultFrom", omission.from());
    }
    if (!scalarLeaf(component.getGenericType())) {
      return;
    }
    var entry = CatalogFieldMetadataSupport.fieldEntry(component, optional);
    constraints(value, entry.shape());
    if (!entry.enumValues().isEmpty()) {
      ObjectNode scalar = value;
      while (scalar.has("items")) {
        scalar = scalar.withObject("items");
      }
      var candidates = scalar.putArray("enum");
      entry.enumValues().forEach(candidates::add);
    }
  }

  private static boolean scalarLeaf(Type type) {
    if (type instanceof ParameterizedType parameterized) {
      return scalarLeaf(parameterized.getActualTypeArguments()[0]);
    }
    return type instanceof Class<?> raw
        && (raw == String.class
            || raw == Path.class
            || raw.isPrimitive()
            || raw.isEnum()
            || Number.class.isAssignableFrom(raw)
            || raw == Boolean.class
            || WireScalar.isValueType(raw)
            || TemporalAccessor.class.isAssignableFrom(raw));
  }

  private void constraints(ObjectNode value, FieldShape shape) {
    if (shape instanceof FieldShape.ListShape list) {
      constraints(value.withObject("items"), list.elementShape());
      return;
    }
    if (!(shape instanceof FieldShape.Scalar scalar)) {
      return;
    }
    for (FieldConstraint constraint : scalar.constraints()) {
      switch (constraint) {
        case FieldConstraint.NonBlank _ -> pattern(value, NON_BLANK_PATTERN);
        case FieldConstraint.StringPattern match ->
            pattern(value, "^(?:" + match.pattern() + ")(?![\\s\\S])");
        case FieldConstraint.LengthRange range ->
            value.put("minLength", range.minimum()).put("maxLength", range.maximum());
        case FieldConstraint.NumberRange range ->
            value.put("minimum", range.minimum()).put("maximum", range.maximum());
        case FieldConstraint.Integral _ -> value.put("type", "integer");
        case FieldConstraint.PathSuffix suffix -> pattern(value, suffixPattern(suffix.suffix()));
      }
    }
  }

  private static void pattern(ObjectNode schema, String pattern) {
    schema.withArray("allOf").addObject().put("pattern", pattern);
  }

  private static String nonBlankPattern() {
    StringBuilder whitespace = new StringBuilder("[^");
    java.util.stream.IntStream.rangeClosed(0, Character.MAX_CODE_POINT)
        .filter(Character::isWhitespace)
        .forEach(whitespace::appendCodePoint);
    return whitespace.append(']').toString();
  }

  private static String suffixPattern(String suffix) {
    StringBuilder pattern = new StringBuilder();
    for (int index = 0; index < suffix.length(); index++) {
      char character = suffix.charAt(index);
      if (character == '.') {
        pattern.append("\\.");
      } else if (Character.isLetter(character)) {
        pattern
            .append('[')
            .append(Character.toLowerCase(character))
            .append(Character.toUpperCase(character))
            .append(']');
      } else {
        pattern.append(character);
      }
    }
    return pattern.append("(?![\\s\\S])").toString();
  }
}
