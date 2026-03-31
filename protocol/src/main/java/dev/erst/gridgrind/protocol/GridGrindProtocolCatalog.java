package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
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
  private static final Set<Class<?>> STRING_FIELD_TYPES =
      Set.of(String.class, java.time.LocalDate.class, java.time.LocalDateTime.class);
  private static final Set<Class<?>> BOOLEAN_FIELD_TYPES = Set.of(boolean.class, Boolean.class);
  private static final Set<Class<?>> NUMERIC_FIELD_TYPES =
      Set.of(
          byte.class,
          short.class,
          int.class,
          long.class,
          float.class,
          double.class,
          Byte.class,
          Short.class,
          Integer.class,
          Long.class,
          Float.class,
          Double.class,
          java.math.BigDecimal.class,
          java.math.BigInteger.class);
  private static final Map<Class<?>, String> NESTED_FIELD_SHAPE_GROUPS =
      Map.ofEntries(
          Map.entry(CellInput.class, "cellInputTypes"),
          Map.entry(HyperlinkTarget.class, "hyperlinkTargetTypes"),
          Map.entry(CellSelection.class, "cellSelectionTypes"),
          Map.entry(SheetSelection.class, "sheetSelectionTypes"),
          Map.entry(NamedRangeSelection.class, "namedRangeSelectionTypes"),
          Map.entry(NamedRangeScope.class, "namedRangeScopeTypes"),
          Map.entry(NamedRangeSelector.class, "namedRangeSelectorTypes"),
          Map.entry(FontHeightInput.class, "fontHeightTypes"));
  private static final Map<Class<?>, String> PLAIN_FIELD_SHAPE_GROUPS =
      Map.ofEntries(
          Map.entry(CommentInput.class, "commentInputType"),
          Map.entry(NamedRangeTarget.class, "namedRangeTargetType"),
          Map.entry(CellStyleInput.class, "cellStyleInputType"),
          Map.entry(CellBorderInput.class, "cellBorderInputType"),
          Map.entry(CellBorderSideInput.class, "cellBorderSideInputType"));
  private static final List<TypeDescriptor> SOURCE_TYPES =
      List.of(
          descriptor(
              GridGrindRequest.WorkbookSource.New.class,
              "NEW",
              "Create a brand-new empty workbook. A new workbook starts with zero sheets;"
                  + " use ENSURE_SHEET to create the first sheet."),
          descriptor(
              GridGrindRequest.WorkbookSource.ExistingFile.class,
              "EXISTING",
              "Open an existing .xlsx workbook from disk."
                  + " Relative paths resolve in the current execution environment."));
  private static final List<TypeDescriptor> PERSISTENCE_TYPES =
      List.of(
          descriptor(
              GridGrindRequest.WorkbookPersistence.None.class,
              "NONE",
              "Keep the workbook in memory only." + " The response persistence.type echoes NONE."),
          descriptor(
              GridGrindRequest.WorkbookPersistence.OverwriteSource.class,
              "OVERWRITE",
              "Overwrite the opened source workbook at source.path."
                  + " No path field is accepted on OVERWRITE;"
                  + " the write target is the same path opened by the EXISTING source."
                  + " Relative source.path values resolve in the current execution environment."
                  + " The response persistence.type echoes OVERWRITE and includes sourcePath"
                  + " (the original source path string) and executionPath (absolute normalized)."),
          descriptor(
              GridGrindRequest.WorkbookPersistence.SaveAs.class,
              "SAVE_AS",
              "Save the workbook to a new .xlsx path."
                  + " Relative paths resolve in the current execution environment."
                  + " The response persistence.type echoes SAVE_AS and includes requestedPath"
                  + " (the literal path from the request) and executionPath (the absolute"
                  + " normalized path where the file was written); they differ when a relative"
                  + " path or a path with .. segments is supplied."
                  + " Missing parent directories are created automatically."));
  private static final List<TypeDescriptor> OPERATION_TYPES =
      List.of(
          descriptor(
              WorkbookOperation.EnsureSheet.class,
              "ENSURE_SHEET",
              "Create the sheet if it does not already exist."
                  + " Sheet names must not exceed 31 characters (Excel limit)."),
          descriptor(
              WorkbookOperation.RenameSheet.class,
              "RENAME_SHEET",
              "Rename an existing sheet."
                  + " The new name must not exceed 31 characters (Excel limit)."),
          descriptor(
              WorkbookOperation.DeleteSheet.class,
              "DELETE_SHEET",
              "Delete an existing sheet."
                  + " A workbook must retain at least one sheet;"
                  + " deleting the last sheet returns INVALID_REQUEST."),
          descriptor(
              WorkbookOperation.MoveSheet.class,
              "MOVE_SHEET",
              "Move a sheet to a zero-based workbook position."
                  + " targetIndex is 0-based: 0 moves the sheet to the front,"
                  + " sheetCount-1 moves it to the back."),
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
              "Set one or more column widths in Excel character units."
                  + " widthCharacters must be > 0 and <= 255 (Excel column width limit)."),
          descriptor(
              WorkbookOperation.SetRowHeight.class,
              "SET_ROW_HEIGHT",
              "Set one or more row heights in Excel point units."
                  + " heightPoints must be > 0 and <= 1638.35 (Excel row height limit: 32767 twips)."),
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
              "Attach a hyperlink to one cell."
                  + " FILE targets use the field name path and normalize file: URIs to plain"
                  + " path strings."),
          descriptor(
              WorkbookOperation.ClearHyperlink.class,
              "CLEAR_HYPERLINK",
              "Remove the hyperlink from one cell; no-op when the cell does not physically exist."),
          descriptor(
              WorkbookOperation.SetComment.class,
              "SET_COMMENT",
              "Attach a plain-text comment to one cell."),
          descriptor(
              WorkbookOperation.ClearComment.class,
              "CLEAR_COMMENT",
              "Remove the comment from one cell; no-op when the cell does not physically exist."),
          descriptor(
              WorkbookOperation.ApplyStyle.class,
              "APPLY_STYLE",
              "Apply a style patch to every cell in a range."
                  + " Write shape: style.border is a nested object"
                  + " { \"all\": { \"style\": \"THIN\" } } or per-side top/right/bottom/left."
                  + " Read shape (GET_CELLS, GET_WINDOW): borders are flat properties"
                  + " topBorderStyle, rightBorderStyle, bottomBorderStyle, leftBorderStyle;"
                  + " the nested border object is write-only."),
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
              "Append a row of typed values after the last value-bearing row."
                  + " Blank rows that carry only style, comment, or hyperlink metadata do not"
                  + " affect the append position."),
          descriptor(
              WorkbookOperation.AutoSizeColumns.class,
              "AUTO_SIZE_COLUMNS",
              "Size columns deterministically from displayed cell content so the resulting"
                  + " widths are stable in headless and container runs."),
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
              "Return workbook-level summary facts including sheet count, sheet names,"
                  + " named range count, and formula recalculation flag."),
          descriptor(
              WorkbookReadOperation.GetNamedRanges.class,
              "GET_NAMED_RANGES",
              "Return named ranges matched by the supplied selection."),
          descriptor(
              WorkbookReadOperation.GetSheetSummary.class,
              "GET_SHEET_SUMMARY",
              "Return structural summary facts for one sheet."
                  + " physicalRowCount is the number of physically materialized rows (sparse)."
                  + " lastRowIndex is the 0-based index of the last materialized row"
                  + " (-1 when empty), including metadata-only rows."
                  + " lastColumnIndex is the 0-based index of the last materialized column"
                  + " in any row (-1 when empty)."),
          descriptor(
              WorkbookReadOperation.GetCells.class,
              "GET_CELLS",
              "Return exact cell snapshots for explicit addresses."
                  + " An invalid or out-of-range address returns INVALID_CELL_ADDRESS, not a blank."
                  + " Each snapshot includes address, declaredType, effectiveType, displayValue,"
                  + " style, and metadata fields. Type-specific fields: stringValue (STRING),"
                  + " numberValue (NUMBER), booleanValue (BOOLEAN), errorValue (ERROR),"
                  + " formula and evaluation (FORMULA). For FORMULA cells, effectiveType is FORMULA"
                  + " and the evaluated result type is in evaluation.effectiveType."
                  + " style.fontHeight is a plain object with both twips and points fields,"
                  + " not the discriminated FontHeightInput write format."),
          descriptor(
              WorkbookReadOperation.GetWindow.class,
              "GET_WINDOW",
              "Return a rectangular window of cell snapshots."
                  + " rowCount * columnCount must not exceed "
                  + WorkbookReadOperation.MAX_WINDOW_CELLS
                  + ". Each cell snapshot has the same shape as GET_CELLS: address,"
                  + " declaredType, effectiveType, displayValue, style, metadata,"
                  + " and type-specific value fields."
                  + " For FORMULA cells, effectiveType is FORMULA and the evaluated result type"
                  + " is in evaluation.effectiveType."
                  + " Response shape: { \"window\": { \"rows\": [ { \"cells\": [...] } ] } }."
                  + " Note: the top-level key is \"window\" and cells are nested under"
                  + " window.rows[N].cells, unlike GET_CELLS which places cells directly"
                  + " under the top-level \"cells\" key."),
          descriptor(
              WorkbookReadOperation.GetMergedRegions.class,
              "GET_MERGED_REGIONS",
              "Return the merged regions defined on one sheet."),
          descriptor(
              WorkbookReadOperation.GetHyperlinks.class,
              "GET_HYPERLINKS",
              "Return hyperlink metadata for selected cells in the same discriminated shape used"
                  + " by SET_HYPERLINK targets. FILE targets are returned as normalized path"
                  + " strings, not file: URIs."),
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
              "Infer a simple schema from a rectangular sheet window."
                  + " rowCount * columnCount must not exceed "
                  + WorkbookReadOperation.MAX_WINDOW_CELLS
                  + "."
                  + " The first row is treated as the header; dataRowCount is 0 when all header"
                  + " cells are blank."
                  + " dominantType is null when all data cells are blank, or when two or more types"
                  + " tie for the highest count."
                  + " Formula cells contribute their evaluated result type (NUMBER, STRING, etc.)"
                  + " to dominantType and observedTypes, not FORMULA."),
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
              "Report hyperlink findings such as malformed, missing, unresolved, or broken"
                  + " targets."
                  + " Relative FILE targets are resolved against the workbook's persisted path"
                  + " when one exists."),
          descriptor(
              WorkbookReadOperation.AnalyzeNamedRangeHealth.class,
              "ANALYZE_NAMED_RANGE_HEALTH",
              "Report named-range findings such as broken references."),
          descriptor(
              WorkbookReadOperation.AnalyzeWorkbookFindings.class,
              "ANALYZE_WORKBOOK_FINDINGS",
              "Run all analysis families (formula health, hyperlink health, named-range health)"
                  + " across the entire workbook and aggregate findings in a single response."
                  + " Equivalent to issuing ANALYZE_FORMULA_HEALTH, ANALYZE_HYPERLINK_HEALTH,"
                  + " and ANALYZE_NAMED_RANGE_HEALTH in one request."));
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
                  descriptor(
                      CellInput.Date.class,
                      "DATE",
                      "Write an ISO-8601 date value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.DateTime.class,
                      "DATE_TIME",
                      "Write an ISO-8601 date-time value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
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
                  descriptor(
                      HyperlinkTarget.File.class,
                      "FILE",
                      "Attach a local or shared file path."
                          + " Accepts plain paths or file: URIs and normalizes to a plain path"
                          + " string."),
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
                      "Specify font height in point units."
                          + " Write format: {\"type\":\"POINTS\",\"points\":13}."
                          + " Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is"
                          + " {\"twips\":260,\"points\":13} with both fields present,"
                          + " not this discriminated type format."),
                  descriptor(
                      FontHeightInput.Twips.class,
                      "TWIPS",
                      "Specify font height in exact twips (20 twips = 1 point)."
                          + " Write format: {\"type\":\"TWIPS\",\"twips\":260}."
                          + " Read-back returns the same plain object shape as POINTS."))));
  private static final List<PlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "commentInputType",
              CommentInput.class,
              "CommentInput",
              "Plain-text comment payload attached to one cell.",
              List.of("visible")),
          plainTypeDescriptor(
              "namedRangeTargetType",
              NamedRangeTarget.class,
              "NamedRangeTarget",
              "Explicit sheet name and A1-style range address for named-range authoring.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleInputType",
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
                  "border")),
          plainTypeDescriptor(
              "cellBorderInputType",
              CellBorderInput.class,
              "CellBorderInput",
              "Border patch for cell styling; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "cellBorderSideInputType",
              CellBorderSideInput.class,
              "CellBorderSideInput",
              "One border side defined by its border style.",
              List.of()));
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

  /**
   * Returns the single catalog entry matching the given operation id, or empty when no entry with
   * that id exists. Searches sourceTypes, persistenceTypes, operationTypes, and readTypes.
   */
  public static java.util.Optional<TypeEntry> entryFor(String id) {
    return allEntries().stream().filter(e -> e.id().equals(id)).findFirst();
  }

  private static List<TypeEntry> allEntries() {
    List<TypeEntry> all = new java.util.ArrayList<>();
    all.addAll(CATALOG.sourceTypes());
    all.addAll(CATALOG.persistenceTypes());
    all.addAll(CATALOG.operationTypes());
    all.addAll(CATALOG.readTypes());
    for (NestedTypeGroup group : CATALOG.nestedTypes()) {
      all.addAll(group.types());
    }
    return all;
  }

  private static Catalog buildCatalog() {
    validateFieldShapeGroupMappings();
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
        PLAIN_TYPE_DESCRIPTORS.stream().map(GridGrindProtocolCatalog::publicPlainGroup).toList());
  }

  private static NestedTypeGroup publicGroup(NestedTypeDescriptor descriptor) {
    return new NestedTypeGroup(descriptor.group(), publicEntries(descriptor.typeDescriptors()));
  }

  private static PlainTypeGroup publicPlainGroup(PlainTypeDescriptor descriptor) {
    return new PlainTypeGroup(descriptor.group(), descriptor.typeEntry());
  }

  private static List<TypeEntry> publicEntries(List<TypeDescriptor> descriptors) {
    return descriptors.stream().map(TypeDescriptor::typeEntry).toList();
  }

  private static NestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<TypeDescriptor> typeDescriptors) {
    return new NestedTypeDescriptor(group, sealedType, typeDescriptors);
  }

  private static PlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return new PlainTypeDescriptor(
        group, recordType, typeEntry(recordType, id, summary, optionalFields));
  }

  private static TypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new TypeDescriptor(
        recordType, typeEntry(recordType, id, summary, List.of(optionalFields)));
  }

  private static TypeEntry typeEntry(
      Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
    return new TypeEntry(id, summary, fieldEntries(recordType, optionalFields));
  }

  private static List<FieldEntry> fieldEntries(
      Class<? extends Record> recordType, List<String> optionalFields) {
    requiredFields(recordType, optionalFields);
    return Arrays.stream(recordType.getRecordComponents())
        .map(
            component ->
                new FieldEntry(
                    component.getName(),
                    optionalFields.contains(component.getName())
                        ? FieldRequirement.OPTIONAL
                        : FieldRequirement.REQUIRED,
                    fieldShape(component.getGenericType()),
                    enumValues(component.getGenericType())))
        .toList();
  }

  /** Returns the machine-readable field shape for one record component type. */
  static FieldShape fieldShape(Type type) {
    Objects.requireNonNull(type, "type must not be null");
    if (type instanceof ParameterizedType parameterizedType) {
      return fieldShape(parameterizedType);
    }
    if (type instanceof Class<?> classType) {
      return fieldShape(classType);
    }
    throw new IllegalStateException("Unsupported catalog field type: " + type);
  }

  /** Returns the machine-readable field shape for one parameterized record component type. */
  static FieldShape fieldShape(ParameterizedType parameterizedType) {
    Objects.requireNonNull(parameterizedType, "parameterizedType must not be null");
    Type rawType = parameterizedType.getRawType();
    if (rawType == List.class) {
      Type[] typeArguments = parameterizedType.getActualTypeArguments();
      if (typeArguments.length != 1) {
        throw new IllegalStateException(
            "List field must declare exactly one type argument: " + parameterizedType);
      }
      return new FieldShape.ListShape(fieldShape(typeArguments[0]));
    }
    throw new IllegalStateException(
        "Unsupported parameterized catalog field type: " + parameterizedType);
  }

  /** Returns the machine-readable field shape for one non-parameterized record component type. */
  static FieldShape fieldShape(Class<?> classType) {
    Objects.requireNonNull(classType, "classType must not be null");
    if (STRING_FIELD_TYPES.contains(classType)) {
      return new FieldShape.Scalar(ScalarType.STRING);
    }
    if (BOOLEAN_FIELD_TYPES.contains(classType)) {
      return new FieldShape.Scalar(ScalarType.BOOLEAN);
    }
    if (isNumericType(classType)) {
      return new FieldShape.Scalar(ScalarType.NUMBER);
    }
    if (classType.isEnum()) {
      return new FieldShape.Scalar(ScalarType.STRING);
    }
    String nestedGroup = NESTED_FIELD_SHAPE_GROUPS.get(classType);
    if (nestedGroup != null) {
      return new FieldShape.NestedTypeGroupRef(nestedGroup);
    }
    String plainGroup = PLAIN_FIELD_SHAPE_GROUPS.get(classType);
    if (plainGroup != null) {
      return new FieldShape.PlainTypeGroupRef(plainGroup);
    }
    throw new IllegalStateException("Unsupported catalog field type: " + classType.getName());
  }

  /** Returns whether one non-parameterized record component type is represented as JSON NUMBER. */
  static boolean isNumericType(Class<?> classType) {
    Objects.requireNonNull(classType, "classType must not be null");
    return NUMERIC_FIELD_TYPES.contains(classType);
  }

  private static void validateFieldShapeGroupMappings() {
    for (NestedTypeDescriptor descriptor : NESTED_TYPE_GROUPS) {
      validateNestedTypeGroupMapping(descriptor.sealedType(), descriptor.group());
    }
    for (PlainTypeDescriptor descriptor : PLAIN_TYPE_DESCRIPTORS) {
      validatePlainTypeGroupMapping(descriptor.recordType(), descriptor.group());
    }
  }

  /** Validates that one nested sealed input type maps to the published field-shape group. */
  static void validateNestedTypeGroupMapping(Class<?> sealedType, String expectedGroup) {
    Objects.requireNonNull(sealedType, "sealedType must not be null");
    String mappedGroup = NESTED_FIELD_SHAPE_GROUPS.get(sealedType);
    if (!expectedGroup.equals(mappedGroup)) {
      throw new IllegalStateException(
          "Field-shape nested group mapping mismatch for "
              + sealedType.getName()
              + ": expected="
              + expectedGroup
              + ", mapped="
              + mappedGroup);
    }
  }

  /** Validates that one plain record input type maps to the published field-shape group. */
  static void validatePlainTypeGroupMapping(Class<?> recordType, String expectedGroup) {
    Objects.requireNonNull(recordType, "recordType must not be null");
    String mappedGroup = PLAIN_FIELD_SHAPE_GROUPS.get(recordType);
    if (!expectedGroup.equals(mappedGroup)) {
      throw new IllegalStateException(
          "Field-shape plain group mapping mismatch for "
              + recordType.getName()
              + ": expected="
              + expectedGroup
              + ", mapped="
              + mappedGroup);
    }
  }

  private static List<String> enumValues(Type type) {
    if (type instanceof Class<?> classType && classType.isEnum()) {
      return Arrays.stream(classType.getEnumConstants())
          .map(value -> ((Enum<?>) value).name())
          .toList();
    }
    return List.of();
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

  /**
   * Builds the failure used when a duplicate entry is detected during catalog construction.
   * Package-private for direct unit testing of the unreachable-in-integration duplicate path.
   */
  @SuppressWarnings("DoNotCallSuggester")
  static IllegalStateException duplicateEntryFailure(String label, Object left, Object right) {
    return new IllegalStateException(
        "Duplicate %s detected while building the protocol catalog: %s / %s"
            .formatted(label, left, right));
  }

  /** Returns the duplicate-entry merge handler used when building ordered protocol maps. */
  static <T> BinaryOperator<T> duplicateEntryHandler(String label) {
    return new DuplicateEntryMerger<>(label);
  }

  /** Merge function that always fails when duplicate ordered-map keys are encountered. */
  record DuplicateEntryMerger<T>(String label) implements BinaryOperator<T> {
    @Override
    public T apply(T left, T right) {
      throw duplicateEntryFailure(label, left, right);
    }
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
  public record TypeEntry(String id, String summary, List<FieldEntry> fields) {
    public TypeEntry {
      id = requireNonBlank(id, "id");
      summary = requireNonBlank(summary, "summary");
      fields = copyFieldEntries(fields, "fields");
    }

    /** Returns the field entry with the given name, or null when this type has no such field. */
    public FieldEntry field(String name) {
      Objects.requireNonNull(name, "name must not be null");
      return fields.stream().filter(field -> field.name().equals(name)).findFirst().orElse(null);
    }
  }

  /** Whether a catalog field is required or optional in the JSON payload. */
  public enum FieldRequirement {
    REQUIRED,
    OPTIONAL
  }

  /** Canonical scalar JSON value types used in the machine-readable field catalog. */
  public enum ScalarType {
    STRING,
    NUMBER,
    BOOLEAN
  }

  /** JSON-serializable field descriptor for one record component. */
  public record FieldEntry(
      String name, FieldRequirement requirement, FieldShape shape, List<String> enumValues) {
    public FieldEntry {
      name = requireNonBlank(name, "name");
      Objects.requireNonNull(requirement, "requirement must not be null");
      Objects.requireNonNull(shape, "shape must not be null");
      enumValues = copyStrings(enumValues, "enumValues");
    }
  }

  /** JSON-serializable recursive field-shape contract used by protocol discovery output. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = FieldShape.Scalar.class, name = "SCALAR"),
    @JsonSubTypes.Type(value = FieldShape.ListShape.class, name = "LIST"),
    @JsonSubTypes.Type(value = FieldShape.NestedTypeGroupRef.class, name = "NESTED_TYPE_GROUP"),
    @JsonSubTypes.Type(value = FieldShape.PlainTypeGroupRef.class, name = "PLAIN_TYPE_GROUP")
  })
  public sealed interface FieldShape
      permits FieldShape.Scalar,
          FieldShape.ListShape,
          FieldShape.NestedTypeGroupRef,
          FieldShape.PlainTypeGroupRef {
    /** Scalar JSON value type such as STRING, NUMBER, or BOOLEAN. */
    record Scalar(ScalarType scalarType) implements FieldShape {
      public Scalar {
        Objects.requireNonNull(scalarType, "scalarType must not be null");
      }
    }

    /** Recursive array shape whose elements all share the same shape. */
    record ListShape(FieldShape elementShape) implements FieldShape {
      public ListShape {
        Objects.requireNonNull(elementShape, "elementShape must not be null");
      }
    }

    /** Reference to one nested discriminated union group published elsewhere in the catalog. */
    record NestedTypeGroupRef(String group) implements FieldShape {
      public NestedTypeGroupRef {
        group = requireNonBlank(group, "group");
      }
    }

    /** Reference to one plain record-type group published elsewhere in the catalog. */
    record PlainTypeGroupRef(String group) implements FieldShape {
      public PlainTypeGroupRef {
        group = requireNonBlank(group, "group");
      }
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

  private static List<FieldEntry> copyFieldEntries(List<FieldEntry> fields, String fieldName) {
    Objects.requireNonNull(fields, fieldName + " must not be null");
    List<FieldEntry> copy = List.copyOf(fields);
    for (int index = 0; index < copy.size(); index++) {
      FieldEntry field =
          Objects.requireNonNull(copy.get(index), fieldName + " must not contain nulls");
      for (int candidateIndex = index + 1; candidateIndex < copy.size(); candidateIndex++) {
        FieldEntry candidate =
            Objects.requireNonNull(copy.get(candidateIndex), fieldName + " must not contain nulls");
        if (field.name().equals(candidate.name())) {
          throw duplicateEntryFailure(fieldName, field, candidate);
        }
      }
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

  private record PlainTypeDescriptor(
      String group, Class<? extends Record> recordType, TypeEntry typeEntry) {
    private PlainTypeDescriptor {
      group = requireNonBlank(group, "group");
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
