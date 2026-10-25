package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * JSON-serializable field descriptor for one record component.
 *
 * <p>{@code defaultBoolean}, when present, is the effective wire default for an optional boolean
 * field. Its absence never implies a default.
 */
public record FieldEntry(
    String name,
    FieldRequirement requirement,
    FieldShape shape,
    List<String> enumValues,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> defaultBoolean,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<EnumValueDocEntry> enumValueDocs,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<String> projectedByFacets,
    @JsonInclude(JsonInclude.Include.NON_DEFAULT) boolean secret) {
  /** Creates one field entry with no optional metadata. */
  public FieldEntry(
      String name, FieldRequirement requirement, FieldShape shape, List<String> enumValues) {
    this(name, requirement, shape, enumValues, Optional.empty(), List.of(), List.of(), false);
  }

  public FieldEntry {
    name = CatalogRecordValidation.requireNonBlank(name, "name");
    Objects.requireNonNull(requirement, "requirement must not be null");
    Objects.requireNonNull(shape, "shape must not be null");
    enumValues = CatalogRecordValidation.copyStrings(enumValues, "enumValues");
    defaultBoolean = Objects.requireNonNullElseGet(defaultBoolean, Optional::empty);
    boolean isBoolean =
        shape instanceof FieldShape.Scalar scalar && scalar.scalarType() == ScalarType.BOOLEAN;
    if (defaultBoolean.isPresent() && (!isBoolean || requirement != FieldRequirement.OPTIONAL)) {
      throw new IllegalArgumentException(
          "defaultBoolean requires an OPTIONAL BOOLEAN field requirement");
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
}
