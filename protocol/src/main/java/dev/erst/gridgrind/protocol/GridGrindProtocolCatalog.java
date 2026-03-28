package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.BinaryOperator;
import java.util.function.Function;
import java.util.stream.Collectors;

/** Publishes the machine-readable GridGrind protocol surface used by CLI discovery commands. */
public final class GridGrindProtocolCatalog {
  private static final String DISCRIMINATOR_FIELD = "type";
  private static final GridGrindRequest REQUEST_TEMPLATE =
      new GridGrindRequest(
          GridGrindProtocolVersion.current(),
          new GridGrindRequest.WorkbookSource.New(),
          new GridGrindRequest.WorkbookPersistence.None(),
          List.of(),
          List.of());
  private static final List<TypeDescriptor> SOURCE_TYPES =
      List.of(
          descriptor(
              GridGrindRequest.WorkbookSource.New.class,
              "NEW",
              "Create a brand-new empty workbook."),
          descriptor(
              GridGrindRequest.WorkbookSource.ExistingFile.class,
              "EXISTING",
              "Open an existing .xlsx workbook from disk."));
  private static final List<TypeDescriptor> PERSISTENCE_TYPES =
      List.of(
          descriptor(
              GridGrindRequest.WorkbookPersistence.None.class,
              "NONE",
              "Keep the workbook in memory only."),
          descriptor(
              GridGrindRequest.WorkbookPersistence.OverwriteSource.class,
              "OVERWRITE",
              "Overwrite the opened source workbook."),
          descriptor(
              GridGrindRequest.WorkbookPersistence.SaveAs.class,
              "SAVE_AS",
              "Save the workbook to a new .xlsx path."));
  private static final List<TypeDescriptor> OPERATION_TYPES =
      List.of(
          descriptor(
              WorkbookOperation.EnsureSheet.class,
              "ENSURE_SHEET",
              "Create the sheet if it does not already exist."),
          descriptor(
              WorkbookOperation.RenameSheet.class, "RENAME_SHEET", "Rename an existing sheet."),
          descriptor(
              WorkbookOperation.DeleteSheet.class, "DELETE_SHEET", "Delete an existing sheet."),
          descriptor(
              WorkbookOperation.MoveSheet.class,
              "MOVE_SHEET",
              "Move a sheet to a zero-based workbook position."),
          descriptor(
              WorkbookOperation.MergeCells.class,
              "MERGE_CELLS",
              "Merge a rectangular A1-style range."),
          descriptor(
              WorkbookOperation.UnmergeCells.class,
              "UNMERGE_CELLS",
              "Remove one merged region by exact range match."),
          descriptor(
              WorkbookOperation.SetColumnWidth.class,
              "SET_COLUMN_WIDTH",
              "Set one or more column widths in Excel character units."),
          descriptor(
              WorkbookOperation.SetRowHeight.class,
              "SET_ROW_HEIGHT",
              "Set one or more row heights in Excel point units."),
          descriptor(
              WorkbookOperation.FreezePanes.class,
              "FREEZE_PANES",
              "Freeze panes using explicit split and visible-origin coordinates."),
          descriptor(
              WorkbookOperation.SetCell.class,
              "SET_CELL",
              "Write one typed value to a single cell."),
          descriptor(
              WorkbookOperation.SetRange.class,
              "SET_RANGE",
              "Write a rectangular grid of typed values."),
          descriptor(
              WorkbookOperation.ClearRange.class,
              "CLEAR_RANGE",
              "Clear value, style, hyperlink, and comment state from a range."),
          descriptor(
              WorkbookOperation.SetHyperlink.class,
              "SET_HYPERLINK",
              "Attach a hyperlink to one cell."),
          descriptor(
              WorkbookOperation.ClearHyperlink.class,
              "CLEAR_HYPERLINK",
              "Remove the hyperlink from one existing cell."),
          descriptor(
              WorkbookOperation.SetComment.class,
              "SET_COMMENT",
              "Attach a plain-text comment to one cell."),
          descriptor(
              WorkbookOperation.ClearComment.class,
              "CLEAR_COMMENT",
              "Remove the comment from one existing cell."),
          descriptor(
              WorkbookOperation.ApplyStyle.class,
              "APPLY_STYLE",
              "Apply a style patch to every cell in a range."),
          descriptor(
              WorkbookOperation.SetNamedRange.class,
              "SET_NAMED_RANGE",
              "Create or replace one workbook- or sheet-scoped named range."),
          descriptor(
              WorkbookOperation.DeleteNamedRange.class,
              "DELETE_NAMED_RANGE",
              "Delete one existing workbook- or sheet-scoped named range."),
          descriptor(
              WorkbookOperation.AppendRow.class,
              "APPEND_ROW",
              "Append a row of typed values after the last populated row."),
          descriptor(
              WorkbookOperation.AutoSizeColumns.class,
              "AUTO_SIZE_COLUMNS",
              "Size populated columns to fit their contents."),
          descriptor(
              WorkbookOperation.EvaluateFormulas.class,
              "EVALUATE_FORMULAS",
              "Evaluate workbook formulas immediately."),
          descriptor(
              WorkbookOperation.ForceFormulaRecalculationOnOpen.class,
              "FORCE_FORMULA_RECALCULATION_ON_OPEN",
              "Mark the workbook to recalculate when opened in Excel."));
  private static final List<TypeDescriptor> READ_TYPES =
      List.of(
          descriptor(
              WorkbookReadOperation.GetWorkbookSummary.class,
              "GET_WORKBOOK_SUMMARY",
              "Return workbook-level summary facts."),
          descriptor(
              WorkbookReadOperation.GetNamedRanges.class,
              "GET_NAMED_RANGES",
              "Return named ranges matched by the supplied selection."),
          descriptor(
              WorkbookReadOperation.GetSheetSummary.class,
              "GET_SHEET_SUMMARY",
              "Return structural summary facts for one sheet."),
          descriptor(
              WorkbookReadOperation.GetCells.class,
              "GET_CELLS",
              "Return exact cell snapshots for explicit addresses."),
          descriptor(
              WorkbookReadOperation.GetWindow.class,
              "GET_WINDOW",
              "Return a rectangular window of cell snapshots."),
          descriptor(
              WorkbookReadOperation.GetMergedRegions.class,
              "GET_MERGED_REGIONS",
              "Return the merged regions defined on one sheet."),
          descriptor(
              WorkbookReadOperation.GetHyperlinks.class,
              "GET_HYPERLINKS",
              "Return hyperlink metadata for selected cells."),
          descriptor(
              WorkbookReadOperation.GetComments.class,
              "GET_COMMENTS",
              "Return comment metadata for selected cells."),
          descriptor(
              WorkbookReadOperation.GetSheetLayout.class,
              "GET_SHEET_LAYOUT",
              "Return freeze-pane, row-height, and column-width metadata."),
          descriptor(
              WorkbookReadOperation.GetFormulaSurface.class,
              "GET_FORMULA_SURFACE",
              "Summarize formula usage patterns across the selected sheets."),
          descriptor(
              WorkbookReadOperation.GetSheetSchema.class,
              "GET_SHEET_SCHEMA",
              "Infer a simple schema from a rectangular sheet window."),
          descriptor(
              WorkbookReadOperation.GetNamedRangeSurface.class,
              "GET_NAMED_RANGE_SURFACE",
              "Summarize the scope and backing kind of named ranges."),
          descriptor(
              WorkbookReadOperation.AnalyzeFormulaHealth.class,
              "ANALYZE_FORMULA_HEALTH",
              "Report formula findings such as errors and volatile usage."),
          descriptor(
              WorkbookReadOperation.AnalyzeHyperlinkHealth.class,
              "ANALYZE_HYPERLINK_HEALTH",
              "Report hyperlink findings such as malformed or broken targets."),
          descriptor(
              WorkbookReadOperation.AnalyzeNamedRangeHealth.class,
              "ANALYZE_NAMED_RANGE_HEALTH",
              "Report named-range findings such as broken references."),
          descriptor(
              WorkbookReadOperation.AnalyzeWorkbookFindings.class,
              "ANALYZE_WORKBOOK_FINDINGS",
              "Aggregate the first analysis family across the workbook."));
  private static final List<NestedTypeDescriptor> NESTED_TYPE_GROUPS =
      List.of(
          nestedTypeGroup(
              "cellInputTypes",
              CellInput.class,
              List.of(
                  descriptor(CellInput.Blank.class, "BLANK", "Write an empty cell."),
                  descriptor(CellInput.Text.class, "TEXT", "Write a string cell value."),
                  descriptor(CellInput.Numeric.class, "NUMBER", "Write a numeric cell value."),
                  descriptor(
                      CellInput.BooleanValue.class, "BOOLEAN", "Write a boolean cell value."),
                  descriptor(CellInput.Date.class, "DATE", "Write an ISO-8601 date value."),
                  descriptor(
                      CellInput.DateTime.class, "DATE_TIME", "Write an ISO-8601 date-time value."),
                  descriptor(
                      CellInput.Formula.class,
                      "FORMULA",
                      "Write an Excel formula. Omit the leading = sign; the engine adds it internally."))),
          nestedTypeGroup(
              "hyperlinkTargetTypes",
              HyperlinkTarget.class,
              List.of(
                  descriptor(HyperlinkTarget.Url.class, "URL", "Attach an absolute URL target."),
                  descriptor(
                      HyperlinkTarget.Email.class,
                      "EMAIL",
                      "Attach an email target without the mailto: prefix."),
                  descriptor(HyperlinkTarget.File.class, "FILE", "Attach a file-system target."),
                  descriptor(
                      HyperlinkTarget.Document.class,
                      "DOCUMENT",
                      "Attach an internal workbook target."))),
          nestedTypeGroup(
              "cellSelectionTypes",
              CellSelection.class,
              List.of(
                  descriptor(
                      CellSelection.AllUsedCells.class,
                      "ALL_USED_CELLS",
                      "Select every physically present cell on the sheet."),
                  descriptor(
                      CellSelection.Selected.class,
                      "SELECTED",
                      "Select only the supplied ordered cell addresses."))),
          nestedTypeGroup(
              "sheetSelectionTypes",
              SheetSelection.class,
              List.of(
                  descriptor(
                      SheetSelection.All.class, "ALL", "Select every sheet in workbook order."),
                  descriptor(
                      SheetSelection.Selected.class,
                      "SELECTED",
                      "Select only the supplied ordered sheet names."))),
          nestedTypeGroup(
              "namedRangeSelectionTypes",
              NamedRangeSelection.class,
              List.of(
                  descriptor(NamedRangeSelection.All.class, "ALL", "Select every named range."),
                  descriptor(
                      NamedRangeSelection.Selected.class,
                      "SELECTED",
                      "Select only the supplied named-range selectors."))),
          nestedTypeGroup(
              "namedRangeScopeTypes",
              NamedRangeScope.class,
              List.of(
                  descriptor(NamedRangeScope.Workbook.class, "WORKBOOK", "Target workbook scope."),
                  descriptor(
                      NamedRangeScope.Sheet.class, "SHEET", "Target one specific sheet scope."))),
          nestedTypeGroup(
              "namedRangeSelectorTypes",
              NamedRangeSelector.class,
              List.of(
                  descriptor(
                      NamedRangeSelector.ByName.class,
                      "BY_NAME",
                      "Match a named range across all scopes by exact name."),
                  descriptor(
                      NamedRangeSelector.WorkbookScope.class,
                      "WORKBOOK_SCOPE",
                      "Match the workbook-scoped named range with the exact name."),
                  descriptor(
                      NamedRangeSelector.SheetScope.class,
                      "SHEET_SCOPE",
                      "Match the sheet-scoped named range on one sheet."))),
          nestedTypeGroup(
              "fontHeightTypes",
              FontHeightInput.class,
              List.of(
                  descriptor(
                      FontHeightInput.Points.class,
                      "POINTS",
                      "Specify font height in point units."),
                  descriptor(
                      FontHeightInput.Twips.class,
                      "TWIPS",
                      "Specify font height in exact twips."))));
  private static final List<PlainTypeGroup> PLAIN_TYPE_GROUPS =
      List.of(
          plainTypeGroup(
              "commentInputType",
              plainDescriptor(
                  CommentInput.class,
                  "CommentInput",
                  "Plain-text comment payload attached to one cell.",
                  List.of("visible"))),
          plainTypeGroup(
              "namedRangeTargetType",
              plainDescriptor(
                  NamedRangeTarget.class,
                  "NamedRangeTarget",
                  "Explicit sheet name and A1-style range address for named-range authoring.",
                  List.of())),
          plainTypeGroup(
              "cellStyleInputType",
              plainDescriptor(
                  CellStyleInput.class,
                  "CellStyleInput",
                  "Style patch applied to a cell or range; at least one field must be set."
                      + " Colors use #RRGGBB hex.",
                  List.of(
                      "numberFormat",
                      "bold",
                      "italic",
                      "wrapText",
                      "horizontalAlignment",
                      "verticalAlignment",
                      "fontName",
                      "fontHeight",
                      "fontColor",
                      "underline",
                      "strikeout",
                      "fillColor",
                      "border"))),
          plainTypeGroup(
              "cellBorderInputType",
              plainDescriptor(
                  CellBorderInput.class,
                  "CellBorderInput",
                  "Border patch for cell styling; at least one side must be set."
                      + " Use 'all' as shorthand for all four sides.",
                  List.of("all", "top", "right", "bottom", "left"))),
          plainTypeGroup(
              "cellBorderSideInputType",
              plainDescriptor(
                  CellBorderSideInput.class,
                  "CellBorderSideInput",
                  "One border side defined by its border style.",
                  List.of())));
  private static final Catalog CATALOG = buildCatalog();

  private GridGrindProtocolCatalog() {}

  /** Returns the minimal successful request emitted by the CLI template command. */
  public static GridGrindRequest requestTemplate() {
    return REQUEST_TEMPLATE;
  }

  /** Returns the machine-readable protocol catalog emitted by the CLI discovery command. */
  public static Catalog catalog() {
    return CATALOG;
  }

  private static Catalog buildCatalog() {
    validateCoverage(GridGrindRequest.WorkbookSource.class, SOURCE_TYPES);
    validateCoverage(GridGrindRequest.WorkbookPersistence.class, PERSISTENCE_TYPES);
    validateCoverage(WorkbookOperation.class, OPERATION_TYPES);
    validateCoverage(WorkbookReadOperation.class, READ_TYPES);
    for (NestedTypeDescriptor nestedTypeGroup : NESTED_TYPE_GROUPS) {
      validateCoverage(nestedTypeGroup.sealedType(), nestedTypeGroup.typeDescriptors());
    }
    return new Catalog(
        GridGrindProtocolVersion.current(),
        DISCRIMINATOR_FIELD,
        publicEntries(SOURCE_TYPES),
        publicEntries(PERSISTENCE_TYPES),
        publicEntries(OPERATION_TYPES),
        publicEntries(READ_TYPES),
        NESTED_TYPE_GROUPS.stream().map(GridGrindProtocolCatalog::publicGroup).toList(),
        List.copyOf(PLAIN_TYPE_GROUPS));
  }

  private static NestedTypeGroup publicGroup(NestedTypeDescriptor descriptor) {
    return new NestedTypeGroup(descriptor.group(), publicEntries(descriptor.typeDescriptors()));
  }

  private static List<TypeEntry> publicEntries(List<TypeDescriptor> descriptors) {
    return descriptors.stream().map(TypeDescriptor::typeEntry).toList();
  }

  private static NestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<TypeDescriptor> typeDescriptors) {
    return new NestedTypeDescriptor(group, sealedType, typeDescriptors);
  }

  private static PlainTypeGroup plainTypeGroup(String group, TypeEntry typeEntry) {
    return new PlainTypeGroup(group, typeEntry);
  }

  private static TypeEntry plainDescriptor(
      Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
    return new TypeEntry(
        id, summary, requiredFields(recordType, optionalFields), List.copyOf(optionalFields));
  }

  private static TypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new TypeDescriptor(
        recordType,
        new TypeEntry(
            id,
            summary,
            requiredFields(recordType, List.of(optionalFields)),
            List.of(optionalFields)));
  }

  /** Returns the required record fields after removing the explicitly optional ones. */
  static List<String> requiredFields(
      Class<? extends Record> recordType, List<String> optionalFields) {
    List<String> recordFields = recordFields(recordType);
    for (String optionalField : optionalFields) {
      if (!recordFields.contains(optionalField)) {
        throw new IllegalStateException(
            "Catalog optional field '%s' does not exist on %s"
                .formatted(optionalField, recordType.getName()));
      }
    }
    return recordFields.stream().filter(field -> !optionalFields.contains(field)).toList();
  }

  private static List<String> recordFields(Class<? extends Record> recordType) {
    return Arrays.stream(recordType.getRecordComponents()).map(RecordComponent::getName).toList();
  }

  private static void validateCoverage(Class<?> sealedType, List<TypeDescriptor> descriptors) {
    validateCoverage(
        sealedType,
        toOrderedMap(
            descriptors,
            TypeDescriptor::recordType,
            descriptor -> descriptor.typeEntry().id(),
            "catalog descriptor"));
  }

  /** Validates that a tagged union and the catalog expose the same ordered {@code type} ids. */
  static void validateCoverage(Class<?> sealedType, Map<Class<?>, String> catalogIds) {
    JsonTypeInfo jsonTypeInfo = sealedType.getAnnotation(JsonTypeInfo.class);
    if (jsonTypeInfo == null || !DISCRIMINATOR_FIELD.equals(jsonTypeInfo.property())) {
      throw new IllegalStateException(
          "Catalog coverage requires %s to use discriminator field '%s'"
              .formatted(sealedType.getName(), DISCRIMINATOR_FIELD));
    }
    JsonSubTypes jsonSubTypes = sealedType.getAnnotation(JsonSubTypes.class);
    if (jsonSubTypes == null) {
      throw new IllegalStateException(
          "Catalog coverage requires @JsonSubTypes on " + sealedType.getName());
    }

    Map<Class<?>, String> annotationIds =
        toOrderedMap(
            Arrays.asList(jsonSubTypes.value()),
            JsonSubTypes.Type::value,
            JsonSubTypes.Type::name,
            "annotation subtype");

    for (Class<?> recordType : catalogIds.keySet()) {
      if (!recordType.isRecord()) {
        throw new IllegalStateException(
            "Catalog entry %s does not target a record type".formatted(recordType));
      }
    }

    if (!annotationIds.keySet().equals(catalogIds.keySet())) {
      throw new IllegalStateException(
          "Catalog coverage mismatch for "
              + sealedType.getName()
              + ": annotated="
              + annotationIds.keySet()
              + ", catalog="
              + catalogIds.keySet());
    }

    for (Map.Entry<Class<?>, String> annotationEntry : annotationIds.entrySet()) {
      String catalogId = catalogIds.get(annotationEntry.getKey());
      if (!annotationEntry.getValue().equals(catalogId)) {
        throw new IllegalStateException(
            "Catalog id mismatch for "
                + annotationEntry.getKey().getName()
                + ": annotation="
                + annotationEntry.getValue()
                + ", catalog="
                + catalogId);
      }
    }
  }

  private static <T, K, V> Map<K, V> toOrderedMap(
      List<T> values, Function<T, K> keyExtractor, Function<T, V> valueExtractor, String label) {
    return values.stream()
        .collect(
            Collectors.toMap(
                keyExtractor, valueExtractor, duplicateEntryHandler(label), LinkedHashMap::new));
  }

  private static <T> BinaryOperator<T> duplicateEntryHandler(String label) {
    return (left, right) -> duplicateEntry(label, left, right);
  }

  private static <T> T duplicateEntry(String label, T left, T right) {
    throw new IllegalStateException(
        "Duplicate %s detected while building the protocol catalog: %s / %s"
            .formatted(label, left, right));
  }

  /** JSON-serializable top-level catalog emitted by {@code --print-protocol-catalog}. */
  public record Catalog(
      GridGrindProtocolVersion protocolVersion,
      String discriminatorField,
      List<TypeEntry> sourceTypes,
      List<TypeEntry> persistenceTypes,
      List<TypeEntry> operationTypes,
      List<TypeEntry> readTypes,
      List<NestedTypeGroup> nestedTypes,
      List<PlainTypeGroup> plainTypes) {
    public Catalog {
      protocolVersion =
          protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
      discriminatorField = requireNonBlank(discriminatorField, "discriminatorField");
      sourceTypes = copyEntries(sourceTypes, "sourceTypes");
      persistenceTypes = copyEntries(persistenceTypes, "persistenceTypes");
      operationTypes = copyEntries(operationTypes, "operationTypes");
      readTypes = copyEntries(readTypes, "readTypes");
      nestedTypes = copyGroups(nestedTypes, "nestedTypes");
      plainTypes = copyPlainGroups(plainTypes, "plainTypes");
    }
  }

  /** JSON-serializable type entry describing one request or nested union variant. */
  public record TypeEntry(
      String id, String summary, List<String> requiredFields, List<String> optionalFields) {
    public TypeEntry {
      id = requireNonBlank(id, "id");
      summary = requireNonBlank(summary, "summary");
      requiredFields = copyStrings(requiredFields, "requiredFields");
      optionalFields = copyStrings(optionalFields, "optionalFields");
    }
  }

  /** JSON-serializable named group of nested tagged union variants. */
  public record NestedTypeGroup(String group, List<TypeEntry> types) {
    public NestedTypeGroup {
      group = requireNonBlank(group, "group");
      types = copyEntries(types, "types");
    }
  }

  /** JSON-serializable named group describing one plain record type with its field shape. */
  public record PlainTypeGroup(String group, TypeEntry type) {
    public PlainTypeGroup {
      group = requireNonBlank(group, "group");
      Objects.requireNonNull(type, "type must not be null");
    }
  }

  private static List<TypeEntry> copyEntries(List<TypeEntry> entries, String fieldName) {
    Objects.requireNonNull(entries, fieldName + " must not be null");
    List<TypeEntry> copy = List.copyOf(entries);
    for (TypeEntry entry : copy) {
      Objects.requireNonNull(entry, fieldName + " must not contain nulls");
    }
    return copy;
  }

  private static List<NestedTypeGroup> copyGroups(List<NestedTypeGroup> groups, String fieldName) {
    Objects.requireNonNull(groups, fieldName + " must not be null");
    List<NestedTypeGroup> copy = List.copyOf(groups);
    for (NestedTypeGroup group : copy) {
      Objects.requireNonNull(group, fieldName + " must not contain nulls");
    }
    return copy;
  }

  private static List<PlainTypeGroup> copyPlainGroups(
      List<PlainTypeGroup> groups, String fieldName) {
    Objects.requireNonNull(groups, fieldName + " must not be null");
    List<PlainTypeGroup> copy = List.copyOf(groups);
    for (PlainTypeGroup group : copy) {
      Objects.requireNonNull(group, fieldName + " must not contain nulls");
    }
    return copy;
  }

  private static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      requireNonBlank(value, fieldName);
    }
    return copy;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private record TypeDescriptor(Class<? extends Record> recordType, TypeEntry typeEntry) {
    private TypeDescriptor {
      Objects.requireNonNull(recordType, "recordType must not be null");
      Objects.requireNonNull(typeEntry, "typeEntry must not be null");
    }
  }

  private record NestedTypeDescriptor(
      String group, Class<?> sealedType, List<TypeDescriptor> typeDescriptors) {
    private NestedTypeDescriptor {
      group = requireNonBlank(group, "group");
      Objects.requireNonNull(sealedType, "sealedType must not be null");
      typeDescriptors = List.copyOf(typeDescriptors);
      for (TypeDescriptor typeDescriptor : typeDescriptors) {
        Objects.requireNonNull(typeDescriptor, "typeDescriptors must not contain nulls");
      }
    }
  }
}
