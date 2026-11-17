package dev.erst.gridgrind.core.catalog;

import java.util.List;
import java.util.stream.Stream;

/** Maps closed catalog lookup values to their published lookup behavior. */
final class CatalogLookupValueSupport {
  private CatalogLookupValueSupport() {}

  static String kind(CatalogLookupValue lookupValue) {
    return switch (lookupValue) {
      case TypeEntry _ -> "ENTRY";
      case NestedTypeGroup _ -> "NESTED_GROUP";
      case PlainTypeGroup _ -> "PLAIN_GROUP";
      case TopLevelTypeGroup _ -> "TOP_LEVEL_GROUP";
    };
  }

  static String summary(CatalogLookupValue lookupValue) {
    return switch (lookupValue) {
      case TypeEntry entryValue -> entryValue.summary();
      case NestedTypeGroup nestedGroupValue ->
          "Nested type group with discriminator "
              + nestedGroupValue.discriminatorField()
              + " and "
              + nestedGroupValue.types().size()
              + " variants.";
      case PlainTypeGroup plainGroupValue -> plainGroupValue.type().summary();
      case TopLevelTypeGroup topLevelGroupValue ->
          "Top-level type category with " + topLevelGroupValue.types().size() + " entries.";
    };
  }

  static String searchableText(
      Catalog catalog, CatalogLookupValue lookupValue, boolean includeReferencedShapes) {
    return switch (lookupValue) {
      case TypeEntry entryValue ->
          entrySearchableText(catalog, entryValue, includeReferencedShapes);
      case NestedTypeGroup nestedGroupValue -> nestedGroupSearchableText(catalog, nestedGroupValue);
      case PlainTypeGroup plainGroupValue -> plainGroupSearchableText(catalog, plainGroupValue);
      case TopLevelTypeGroup topLevelGroupValue ->
          topLevelGroupSearchableText(catalog, topLevelGroupValue.types());
    };
  }

  static List<String> relatedEntryIds(Catalog catalog, CatalogLookupValue lookupValue) {
    return switch (lookupValue) {
      case TypeEntry _ -> List.of();
      case NestedTypeGroup nestedGroupValue ->
          CatalogShapeTextSupport.relatedEntryIdsForFieldShape(
              catalog, new FieldShape.NestedTypeGroupRef(nestedGroupValue.group()));
      case PlainTypeGroup plainGroupValue ->
          CatalogShapeTextSupport.relatedEntryIdsForFieldShape(
              catalog, new FieldShape.PlainTypeGroupRef(plainGroupValue.group()));
      case TopLevelTypeGroup _ -> List.of();
    };
  }

  private static String entrySearchableText(
      Catalog catalog, TypeEntry entry, boolean includeReferencedShapes) {
    return String.join(
            " ",
            entry.fields().stream().flatMap(CatalogLookupValueSupport::fieldSearchTokens).toList())
        + " "
        + String.join(
            " ",
            entry.targetSelectors().stream()
                .flatMap(selector ->
                    Stream.concat(Stream.of(selector.family()), selector.typeIds().stream()))
                .toList())
        + (includeReferencedShapes
            ? " " + CatalogShapeTextSupport.referencedShapeText(catalog, entry.fields())
            : "")
        + " "
        + CatalogNoteResolutionSupport.referencedNoteText(catalog, entry.noteRefs())
        + " "
        + entry
            .stepTemplate()
            .map(template -> template.bodyField() + " " + String.join(" ", template.notes()))
            .orElse("");
  }

  private static String nestedGroupSearchableText(Catalog catalog, NestedTypeGroup group) {
    return group.types().stream()
        .flatMap(entry ->
            Stream.concat(
                Stream.of(
                    entry.id(),
                    entry.summary(),
                    CatalogNoteResolutionSupport.referencedNoteText(catalog, entry.noteRefs())),
                entry.fields().stream().flatMap(CatalogLookupValueSupport::fieldSearchTokens)))
        .collect(java.util.stream.Collectors.joining(" "));
  }

  private static String plainGroupSearchableText(Catalog catalog, PlainTypeGroup group) {
    TypeEntry entry = group.type();
    return Stream.concat(
            Stream.of(
                entry.id(),
                entry.summary(),
                CatalogNoteResolutionSupport.referencedNoteText(catalog, entry.noteRefs())),
            entry.fields().stream().flatMap(CatalogLookupValueSupport::fieldSearchTokens))
        .collect(java.util.stream.Collectors.joining(" "));
  }

  private static String topLevelGroupSearchableText(Catalog catalog, List<TypeEntry> types) {
    return types.stream()
        .flatMap(entry ->
            Stream.concat(
                Stream.of(
                    entry.id(),
                    entry.summary(),
                    CatalogNoteResolutionSupport.referencedNoteText(catalog, entry.noteRefs())),
                entry.fields().stream().flatMap(CatalogLookupValueSupport::fieldSearchTokens)))
        .collect(java.util.stream.Collectors.joining(" "));
  }

  private static Stream<String> fieldSearchTokens(FieldEntry field) {
    return Stream.concat(
        Stream.of(field.name()),
        Stream.concat(
            field.enumValues().stream(),
            Stream.concat(
                field.projectedByFacets().stream(),
                field.enumValueDocs().stream()
                    .flatMap(doc -> Stream.of(doc.value(), doc.summary())))));
  }
}
