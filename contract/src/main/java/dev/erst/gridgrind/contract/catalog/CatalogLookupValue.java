package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;
import java.util.stream.Stream;

/** Internal typed carrier for one published lookup surface. */
abstract sealed class CatalogLookupValue
    permits CatalogLookupValue.EntryLookupValue,
        CatalogLookupValue.NestedGroupLookupValue,
        CatalogLookupValue.PlainGroupLookupValue,
        CatalogLookupValue.TopLevelGroupLookupValue {
  abstract Object rawValue();

  abstract String kind();

  abstract String summary();

  abstract String searchableText(Catalog catalog, boolean includeReferencedShapes);

  abstract List<String> relatedEntryIds(Catalog catalog);

  /** Lookup wrapper around one published type entry. */
  static final class EntryLookupValue extends CatalogLookupValue {
    private final TypeEntry entry;

    EntryLookupValue(TypeEntry entry) {
      this.entry = Objects.requireNonNull(entry, "entry must not be null");
    }

    TypeEntry entry() {
      return entry;
    }

    @Override
    Object rawValue() {
      return entry;
    }

    @Override
    String kind() {
      return "ENTRY";
    }

    @Override
    String summary() {
      return entry.summary();
    }

    @Override
    String searchableText(Catalog catalog, boolean includeReferencedShapes) {
      return String.join(
              " ",
              entry.fields().stream()
                  .flatMap(
                      field -> Stream.concat(Stream.of(field.name()), field.enumValues().stream()))
                  .toList())
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
          + entry
              .stepTemplate()
              .map(template -> template.bodyField() + " " + String.join(" ", template.notes()))
              .orElse("");
    }

    @Override
    List<String> relatedEntryIds(Catalog catalog) {
      return List.of();
    }
  }

  /** Lookup wrapper around one published nested type group. */
  static final class NestedGroupLookupValue extends CatalogLookupValue {
    private final NestedTypeGroup group;

    NestedGroupLookupValue(NestedTypeGroup group) {
      this.group = Objects.requireNonNull(group, "group must not be null");
    }

    @Override
    Object rawValue() {
      return group;
    }

    @Override
    String kind() {
      return "NESTED_GROUP";
    }

    @Override
    String summary() {
      return "Nested type group with discriminator "
          + group.discriminatorField()
          + " and "
          + group.types().size()
          + " variants.";
    }

    @Override
    String searchableText(Catalog catalog, boolean includeReferencedShapes) {
      return group.types().stream()
          .flatMap(
              entry ->
                  Stream.concat(
                      Stream.of(entry.id(), entry.summary()),
                      entry.fields().stream().map(FieldEntry::name)))
          .collect(java.util.stream.Collectors.joining(" "));
    }

    @Override
    List<String> relatedEntryIds(Catalog catalog) {
      return CatalogShapeTextSupport.relatedEntryIdsForFieldShape(
          catalog, new FieldShape.NestedTypeGroupRef(group.group()));
    }
  }

  /** Lookup wrapper around one published plain type group. */
  static final class PlainGroupLookupValue extends CatalogLookupValue {
    private final PlainTypeGroup group;

    PlainGroupLookupValue(PlainTypeGroup group) {
      this.group = Objects.requireNonNull(group, "group must not be null");
    }

    @Override
    Object rawValue() {
      return group;
    }

    @Override
    String kind() {
      return "PLAIN_GROUP";
    }

    @Override
    String summary() {
      return group.type().summary();
    }

    @Override
    String searchableText(Catalog catalog, boolean includeReferencedShapes) {
      return Stream.concat(
              Stream.of(group.type().id(), group.type().summary()),
              group.type().fields().stream().map(FieldEntry::name))
          .collect(java.util.stream.Collectors.joining(" "));
    }

    @Override
    List<String> relatedEntryIds(Catalog catalog) {
      return CatalogShapeTextSupport.relatedEntryIdsForFieldShape(
          catalog, new FieldShape.PlainTypeGroupRef(group.group()));
    }
  }

  /** Lookup wrapper around one top-level operation category (e.g. mutationActionTypes). */
  static final class TopLevelGroupLookupValue extends CatalogLookupValue {
    private final String group;
    private final List<TypeEntry> types;

    TopLevelGroupLookupValue(String group, List<TypeEntry> types) {
      this.group = Objects.requireNonNull(group, "group must not be null");
      this.types = List.copyOf(Objects.requireNonNull(types, "types must not be null"));
    }

    @Override
    Object rawValue() {
      return new TopLevelTypeGroup(group, types);
    }

    @Override
    String kind() {
      return "TOP_LEVEL_GROUP";
    }

    @Override
    String summary() {
      return "Top-level type category with " + types.size() + " entries.";
    }

    @Override
    String searchableText(Catalog catalog, boolean includeReferencedShapes) {
      return types.stream()
          .flatMap(
              entry ->
                  Stream.concat(
                      Stream.of(entry.id(), entry.summary()),
                      entry.fields().stream().map(FieldEntry::name)))
          .collect(java.util.stream.Collectors.joining(" "));
    }

    @Override
    List<String> relatedEntryIds(Catalog catalog) {
      return List.of();
    }
  }
}
