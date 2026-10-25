package dev.erst.gridgrind.contract.catalog;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/** Shared note-resolution helpers for catalog payloads, search, and validation. */
final class CatalogNoteResolutionSupport {
  private CatalogNoteResolutionSupport() {}

  static List<CatalogNote> referencedNotes(Catalog catalog, Object value) {
    Objects.requireNonNull(catalog, "catalog must not be null");
    Objects.requireNonNull(value, "value must not be null");
    return referencedNotes(catalog, noteRefsFor(value));
  }

  static String referencedNoteText(Catalog catalog, List<String> noteRefs) {
    return referencedNotes(catalog, noteRefs).stream()
        .map(CatalogNote::text)
        .collect(Collectors.joining(" "));
  }

  static void validateCatalogNoteRefs(
      List<List<TypeEntry>> topLevelEntryGroups,
      List<NestedTypeGroup> nestedTypes,
      List<PlainTypeGroup> plainTypes,
      List<CatalogNote> notes) {
    Objects.requireNonNull(topLevelEntryGroups, "topLevelEntryGroups must not be null");
    Map<String, CatalogNote> notesById =
        notes.stream().collect(Collectors.toUnmodifiableMap(CatalogNote::id, note -> note));
    for (List<TypeEntry> entries : topLevelEntryGroups) {
      validateEntryListNoteRefs(entries, notesById);
    }
    for (NestedTypeGroup group : nestedTypes) {
      validateEntryListNoteRefs(group.types(), notesById);
    }
    for (PlainTypeGroup group : plainTypes) {
      validateEntryNoteRefs(group.type(), notesById);
    }
  }

  private static void validateEntryListNoteRefs(
      List<TypeEntry> entries, Map<String, CatalogNote> notesById) {
    for (TypeEntry entry : entries) {
      validateEntryNoteRefs(entry, notesById);
    }
  }

  private static void validateEntryNoteRefs(TypeEntry entry, Map<String, CatalogNote> notesById) {
    for (String noteRef : entry.noteRefs()) {
      if (!notesById.containsKey(noteRef)) {
        throw new IllegalStateException(
            "Catalog entry " + entry.id() + " references unknown note id " + noteRef);
      }
    }
  }

  private static List<CatalogNote> referencedNotes(Catalog catalog, List<String> noteRefs) {
    Objects.requireNonNull(catalog, "catalog must not be null");
    Objects.requireNonNull(noteRefs, "noteRefs must not be null");
    Set<String> uniqueNoteIds = new LinkedHashSet<>(noteRefs);
    return uniqueNoteIds.stream()
        .map(
            noteId ->
                catalog
                    .note(noteId)
                    .orElseThrow(
                        () ->
                            new IllegalStateException(
                                "Catalog note id " + noteId + " is not published")))
        .toList();
  }

  private static List<String> noteRefsFor(Object value) {
    Set<String> noteRefs = new LinkedHashSet<>();
    collectNoteRefs(value, noteRefs::add);
    return List.copyOf(noteRefs);
  }

  private static void collectNoteRefs(Object value, Consumer<String> sink) {
    Objects.requireNonNull(value, "value must not be null");
    Objects.requireNonNull(sink, "sink must not be null");
    if (value instanceof TypeEntry entry) {
      entry.noteRefs().forEach(sink);
      return;
    }
    if (value instanceof NestedTypeGroup group) {
      group.types().forEach(entry -> entry.noteRefs().forEach(sink));
      return;
    }
    if (value instanceof PlainTypeGroup group) {
      group.type().noteRefs().forEach(sink);
      return;
    }
    if (value instanceof TopLevelTypeGroup group) {
      group.types().forEach(entry -> entry.noteRefs().forEach(sink));
    }
  }
}
