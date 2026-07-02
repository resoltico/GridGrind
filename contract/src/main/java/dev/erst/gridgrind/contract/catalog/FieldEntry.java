package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.List;
import java.util.Objects;

/** JSON-serializable field descriptor for one record component. */
public record FieldEntry(
    String name,
    FieldRequirement requirement,
    FieldShape shape,
    List<String> enumValues,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<EnumValueDocEntry> enumValueDocs,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<String> projectedByFacets) {
  /** Creates one field entry with no projected-facet gating metadata. */
  public FieldEntry(
      String name, FieldRequirement requirement, FieldShape shape, List<String> enumValues) {
    this(name, requirement, shape, enumValues, List.of(), List.of());
  }

  public FieldEntry {
    name = CatalogRecordValidation.requireNonBlank(name, "name");
    Objects.requireNonNull(requirement, "requirement must not be null");
    Objects.requireNonNull(shape, "shape must not be null");
    enumValues = CatalogRecordValidation.copyStrings(enumValues, "enumValues");
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
