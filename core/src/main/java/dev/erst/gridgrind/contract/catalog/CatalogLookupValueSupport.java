package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.stream.Stream;

/** Maps closed catalog lookup values to their published lookup behavior. */
final class CatalogLookupValueSupport {
  private CatalogLookupValueSupport() {}

  static Object rawValue(CatalogLookupValue lookupValue) {
    return switch (lookupValue) {
      case CatalogEntryLookupValue entryValue -> entryValue.entry();
      case CatalogNestedGroupLookupValue nestedGroupValue -> nestedGroupValue.group();
      case CatalogPlainGroupLookupValue plainGroupValue -> plainGroupValue.group();
      case CatalogTopLevelGroupLookupValue topLevelGroupValue ->
          new TopLevelTypeGroup(topLevelGroupValue.group(), topLevelGroupValue.types());
    };
  }

  static String kind(CatalogLookupValue lookupValue) {
    return switch (lookupValue) {
      case CatalogEntryLookupValue _ -> "ENTRY";
      case CatalogNestedGroupLookupValue _ -> "NESTED_GROUP";
      case CatalogPlainGroupLookupValue _ -> "PLAIN_GROUP";
      case CatalogTopLevelGroupLookupValue _ -> "TOP_LEVEL_GROUP";
    };
  }

  static String summary(CatalogLookupValue lookupValue) {
    return switch (lookupValue) {
      case CatalogEntryLookupValue entryValue -> entryValue.entry().summary();
      case CatalogNestedGroupLookupValue nestedGroupValue ->
          "Nested type group with discriminator "
              + nestedGroupValue.group().discriminatorField()
              + " and "
              + nestedGroupValue.group().types().size()
              + " variants.";
      case CatalogPlainGroupLookupValue plainGroupValue -> plainGroupValue.group().type().summary();
      case CatalogTopLevelGroupLookupValue topLevelGroupValue ->
          "Top-level type category with " + topLevelGroupValue.types().size() + " entries.";
    };
  }

  static String searchableText(
      Catalog catalog, CatalogLookupValue lookupValue, boolean includeReferencedShapes) {
    return switch (lookupValue) {
      case CatalogEntryLookupValue entryValue ->
          entrySearchableText(catalog, entryValue.entry(), includeReferencedShapes);
      case CatalogNestedGroupLookupValue nestedGroupValue ->
          nestedGroupSearchableText(catalog, nestedGroupValue.group());
      case CatalogPlainGroupLookupValue plainGroupValue ->
          plainGroupSearchableText(catalog, plainGroupValue.group());
      case CatalogTopLevelGroupLookupValue topLevelGroupValue ->
          topLevelGroupSearchableText(catalog, topLevelGroupValue.types());
    };
  }

  static List<String> relatedEntryIds(Catalog catalog, CatalogLookupValue lookupValue) {
    return switch (lookupValue) {
      case CatalogEntryLookupValue _ -> List.of();
      case CatalogNestedGroupLookupValue nestedGroupValue ->
          CatalogShapeTextSupport.relatedEntryIdsForFieldShape(
              catalog, new FieldShape.NestedTypeGroupRef(nestedGroupValue.group().group()));
      case CatalogPlainGroupLookupValue plainGroupValue ->
          CatalogShapeTextSupport.relatedEntryIdsForFieldShape(
              catalog, new FieldShape.PlainTypeGroupRef(plainGroupValue.group().group()));
      case CatalogTopLevelGroupLookupValue _ -> List.of();
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
                .flatMap(
                    selector ->
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
        .flatMap(
            entry ->
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
        .flatMap(
            entry ->
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
