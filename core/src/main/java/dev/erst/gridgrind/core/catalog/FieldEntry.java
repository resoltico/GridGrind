package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * JSON-serializable field descriptor for one record component.
 *
 * <p>A declared default is either an immutable JSON value tree or a sibling component reference.
 * Absence of both fields never implies a default.
 */
public record FieldEntry(
    String name,
    FieldRequirement requirement,
    FieldShape shape,
    List<String> enumValues,
    Optional<Object> defaultValue,
    Optional<String> defaultFrom,
    @WireDefault(value = "[]", omit = true) List<EnumValueDocEntry> enumValueDocs,
    @WireDefault(value = "[]", omit = true) List<String> projectedByFacets,
    @WireDefault(value = "false", omit = true) boolean secret) {
  /** Creates one field entry with no optional metadata. */
  public FieldEntry(
      String name, FieldRequirement requirement, FieldShape shape, List<String> enumValues) {
    this(
        name,
        requirement,
        shape,
        enumValues,
        Optional.empty(),
        Optional.empty(),
        List.of(),
        List.of(),
        false);
  }

  public FieldEntry {
    name = TextValidation.requireNonBlank(name, "name");
    Objects.requireNonNull(requirement, "requirement must not be null");
    Objects.requireNonNull(shape, "shape must not be null");
    enumValues = CatalogRecordValidation.copyStrings(enumValues, "enumValues");
    Objects.requireNonNull(defaultValue, "defaultValue must not be null");
    Objects.requireNonNull(defaultFrom, "defaultFrom must not be null");
    if ((defaultValue.isPresent() || defaultFrom.isPresent())
        && requirement != FieldRequirement.OPTIONAL) {
      throw new IllegalArgumentException("Defaults require an OPTIONAL field requirement");
    }
    if (defaultValue.isPresent() && defaultFrom.isPresent()) {
      throw new IllegalArgumentException("A default cannot be both constant and derived");
    }
    defaultValue = defaultValue.map(dev.erst.gridgrind.core.metadata.ModelDefaults::wireValue);
    if (defaultValue.isPresent() && !matches(shape, defaultValue.orElseThrow())) {
      throw new IllegalArgumentException("defaultValue must match the field shape");
    }
    enumValueDocs = Objects.requireNonNullElseGet(enumValueDocs, List::of);
    enumValueDocs = CatalogRecordValidation.copyEnumValueDocs(enumValueDocs, "enumValueDocs");
    if (!enumValueDocs.isEmpty()
        && !enumValues.equals(enumValueDocs.stream().map(EnumValueDocEntry::value).toList())) {
      throw new IllegalArgumentException(
          "enumValueDocs must document every enumValues entry in published order");
    }
    projectedByFacets = Objects.requireNonNullElseGet(projectedByFacets, List::of);
    projectedByFacets =
        CatalogRecordValidation.copyUniqueStrings(projectedByFacets, "projectedByFacets");
    if (!projectedByFacets.isEmpty() && requirement != FieldRequirement.OPTIONAL) {
      throw new IllegalArgumentException(
          "projectedByFacets require the field requirement to be OPTIONAL");
    }
  }

  private static boolean matches(FieldShape shape, Object value) {
    return switch (shape) {
      case FieldShape.Scalar scalar ->
          switch (scalar.scalarType()) {
            case STRING -> value instanceof String;
            case NUMBER -> value instanceof Number;
            case BOOLEAN -> value instanceof Boolean;
          };
      case FieldShape.ListShape list ->
          value instanceof List<?> values
              && values.stream().allMatch(item -> matches(list.elementShape(), item));
      default -> value instanceof java.util.Map<?, ?>;
    };
  }
}
