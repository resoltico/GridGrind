package dev.erst.gridgrind.contract.catalog;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Stream;

/** OOXML shape text rendering and shape-reference traversal for catalog search indexing. */
final class CatalogShapeTextSupport {
  private CatalogShapeTextSupport() {}

  static String referencedShapeText(Catalog catalog, List<FieldEntry> fields) {
    return fields.stream()
        .map(FieldEntry::shape)
        .map(shape -> referencedShapeText(catalog, shape))
        .filter(text -> !text.isBlank())
        .collect(java.util.stream.Collectors.joining(" "));
  }

  private static String referencedShapeText(Catalog catalog, FieldShape shape) {
    return switch (shape) {
      case FieldShape.Scalar _ -> "";
      case FieldShape.ListShape listShape -> referencedShapeText(catalog, listShape.elementShape());
      case FieldShape.TopLevelTypeSetRef topLevelTypeSetRef ->
          topLevelTypeText(catalog, topLevelTypeSetRef.typeSet());
      case FieldShape.NestedTypeGroupRef nestedTypeGroupRef ->
          nestedGroupText(catalog, nestedTypeGroupRef.group());
      case FieldShape.NestedTypeGroupUnionRef unionRef ->
          unionRef.groups().stream()
              .map(group -> nestedGroupText(catalog, group))
              .collect(java.util.stream.Collectors.joining(" "));
      case FieldShape.PlainTypeGroupRef plainTypeGroupRef ->
          plainGroupText(catalog, plainTypeGroupRef.group());
    };
  }

  private static String topLevelTypeText(Catalog catalog, String typeSet) {
    return topLevelEntriesForTypeSet(catalog, typeSet).stream()
        .flatMap(entry -> Stream.of(entry.id(), entry.summary()))
        .collect(java.util.stream.Collectors.joining(" "));
  }

  private static String nestedGroupText(Catalog catalog, String groupName) {
    return nestedGroupDescriptor(catalog, groupName).types().stream()
        .flatMap(
            entry ->
                Stream.concat(
                    Stream.of(entry.id(), entry.summary()),
                    entry.fields().stream().map(FieldEntry::name)))
        .collect(java.util.stream.Collectors.joining(" "));
  }

  private static String plainGroupText(Catalog catalog, String groupName) {
    PlainTypeGroup group = plainGroupDescriptor(catalog, groupName);
    return Stream.concat(
            Stream.of(group.type().id(), group.type().summary()),
            group.type().fields().stream().map(FieldEntry::name))
        .collect(java.util.stream.Collectors.joining(" "));
  }

  static List<TypeEntry> topLevelEntriesForTypeSet(Catalog catalog, String typeSet) {
    return switch (typeSet) {
      case "sourceTypes" -> catalog.sourceTypes();
      case "persistenceTypes" -> catalog.persistenceTypes();
      case "stepTypes" -> catalog.stepTypes();
      case "mutationActionTypes" -> catalog.mutationActionTypes();
      case "assertionTypes" -> catalog.assertionTypes();
      case "inspectionQueryTypes" -> catalog.inspectionQueryTypes();
      default -> throw new IllegalArgumentException("Unsupported top-level type set: " + typeSet);
    };
  }

  static NestedTypeGroup nestedGroupDescriptor(Catalog catalog, String groupName) {
    return catalog.nestedTypes().stream()
        .filter(group -> group.group().equals(groupName))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown nested type group: " + groupName));
  }

  static PlainTypeGroup plainGroupDescriptor(Catalog catalog, String groupName) {
    return catalog.plainTypes().stream()
        .filter(group -> group.group().equals(groupName))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown plain type group: " + groupName));
  }

  static List<TypeEntry> nestedGroupEntries(Catalog catalog, String groupName) {
    return nestedGroupDescriptor(catalog, groupName).types();
  }

  static boolean isNestedEntryGroup(Catalog catalog, String catalogGroup) {
    return catalog.nestedTypes().stream().anyMatch(group -> group.group().equals(catalogGroup));
  }

  static boolean isPlainEntryGroup(Catalog catalog, String catalogGroup) {
    return catalog.plainTypes().stream().anyMatch(group -> group.group().equals(catalogGroup));
  }

  static List<String> relatedEntryIdsForFieldShape(Catalog catalog, FieldShape targetShape) {
    return CatalogRefResolutionSupport.topLevelEntryRefs(catalog).stream()
        .filter(
            entryRef ->
                entryReferencesShape(
                    catalog, entryRef.entry().fields(), targetShape, new LinkedHashSet<>()))
        .map(CatalogRefResolutionSupport.CatalogEntryRef::qualifiedId)
        .toList();
  }

  static boolean entryReferencesShape(
      Catalog catalog,
      List<FieldEntry> fields,
      FieldShape targetShape,
      Set<String> recursionGuard) {
    return fields.stream()
        .anyMatch(field -> shapeReferences(catalog, field.shape(), targetShape, recursionGuard));
  }

  static List<String> relatedEntryIdsFor(
      Catalog catalog, CatalogLookupRef ref, CatalogLookupValue lookupValue) {
    if (isNestedEntryGroup(catalog, ref.catalogGroup())) {
      return relatedEntryIdsForFieldShape(
          catalog, new FieldShape.NestedTypeGroupRef(ref.catalogGroup()));
    }
    if (isPlainEntryGroup(catalog, ref.catalogGroup())) {
      return relatedEntryIdsForFieldShape(
          catalog, new FieldShape.PlainTypeGroupRef(ref.catalogGroup()));
    }
    return lookupValue.relatedEntryIds(catalog);
  }

  private static boolean shapeReferences(
      Catalog catalog,
      FieldShape candidateShape,
      FieldShape targetShape,
      Set<String> recursionGuard) {
    if (candidateShape.equals(targetShape)) {
      return true;
    }
    return switch (candidateShape) {
      case FieldShape.Scalar _ -> false;
      case FieldShape.ListShape listShape ->
          shapeReferences(catalog, listShape.elementShape(), targetShape, recursionGuard);
      case FieldShape.TopLevelTypeSetRef _ -> false;
      case FieldShape.NestedTypeGroupRef nestedTypeGroupRef -> {
        String guardKey = "nested:" + nestedTypeGroupRef.group();
        if (!recursionGuard.add(guardKey)) {
          yield false;
        }
        boolean matched =
            nestedGroupEntries(catalog, nestedTypeGroupRef.group()).stream()
                .anyMatch(
                    entry ->
                        entryReferencesShape(catalog, entry.fields(), targetShape, recursionGuard));
        recursionGuard.remove(guardKey);
        yield matched;
      }
      case FieldShape.NestedTypeGroupUnionRef unionRef ->
          unionRef.groups().stream()
              .flatMap(group -> nestedGroupEntries(catalog, group).stream())
              .anyMatch(
                  entry ->
                      entryReferencesShape(catalog, entry.fields(), targetShape, recursionGuard));
      case FieldShape.PlainTypeGroupRef plainTypeGroupRef -> {
        String guardKey = "plain:" + plainTypeGroupRef.group();
        if (!recursionGuard.add(guardKey)) {
          yield false;
        }
        boolean matched =
            entryReferencesShape(
                catalog,
                plainGroupDescriptor(catalog, plainTypeGroupRef.group()).type().fields(),
                targetShape,
                recursionGuard);
        recursionGuard.remove(guardKey);
        yield matched;
      }
    };
  }
}
