package dev.erst.gridgrind.engine.pivots;

import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotTableNaming;
import dev.erst.gridgrind.engine.backend.xssf.PoiPackageInspection;
import dev.erst.gridgrind.engine.backend.xssf.PoiRelationRemoval;
import dev.erst.gridgrind.engine.backend.xssf.SheetNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.names.NamedRangeAdapter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCache;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheRecords;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;

/**
 * Owns pivot creation/deletion, cache/package lifecycle, source binding, factual readback and
 * health.
 */
// The unexported adapter owns one complete feature; method count does not split its boundary.
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired", "PMD.ExcessivePublicCount"})
public final class PivotAdapter {
  private final BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover;

  public PivotAdapter() {
    this(PoiRelationRemoval.defaultRemover());
  }

  public PivotAdapter(BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    this.poiRelationRemover =
        Objects.requireNonNull(poiRelationRemover, "poiRelationRemover must not be null");
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  public void setPivotTable(XssfWorkbook workbook, PivotTableSpec definition) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    Optional<PivotHandle> existing = pivotByName(workbook, definition.name());
    if (existing.isPresent()
        && !existing.orElseThrow().sheetName().equals(definition.sheetName().value())) {
      throw new IllegalArgumentException(
          "pivot table name already exists on a different sheet: " + definition.name());
    }

    ResolvedAuthoringSource source = resolveAuthoringSource(workbook, definition.source());
    SourceColumns columns = sourceColumns(source.sheet(), source.area(), source.description());
    CellReference anchor = new CellReference(definition.anchor().topLeftAddress().value());
    requireSupportedReportFilterAnchor(definition, anchor);

    if (existing.isPresent()) {
      deletePivotHandle(
          workbook, existing.orElseThrow(), allPivotTables(workbook), poiRelationRemover);
    }

    primePivotTableAllocator(workbook.xssfWorkbook(), existing.map(PivotHandle::table));
    try {
      XSSFPivotTable pivotTable = createPivotTable(workbook, definition, source, anchor);
      normalizeCacheId(workbook.xssfWorkbook(), pivotTable);
      pivotTable.getCTPivotTableDefinition().setName(definition.name());
      applyPivotFields(pivotTable, definition, columns);
    } finally {
      rebuildPivotTableRegistry(workbook.xssfWorkbook());
    }
  }

  private void requireSupportedReportFilterAnchor(PivotTableSpec definition, CellReference anchor) {
    if (!definition.reportFilters().isEmpty() && anchor.getRow() < 2) {
      throw new IllegalArgumentException(
          "pivot tables with reportFilters require anchor.topLeftAddress on row 3 or lower because Apache POI moves page filters below the first two rows");
    }
  }

  private void applyPivotFields(
      XSSFPivotTable pivotTable, PivotTableSpec definition, SourceColumns columns) {
    for (String rowLabel : definition.rowLabels()) {
      pivotTable.addRowLabel(columns.relativeIndex(rowLabel));
    }
    for (String columnLabel : definition.columnLabels()) {
      pivotTable.addColLabel(columns.relativeIndex(columnLabel));
    }
    for (String reportFilter : definition.reportFilters()) {
      pivotTable.addReportFilter(columns.relativeIndex(reportFilter));
    }
    for (PivotTableSpec.DataField dataField : definition.dataFields()) {
      pivotTable.addColumnLabel(
          toPoi(dataField.function()),
          columns.relativeIndex(dataField.sourceColumnName()),
          dataField.displayName(),
          dataField.valueFormat().orElse(null));
    }
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet name. */
  public void deletePivotTable(XssfWorkbook workbook, String name, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String validatedName = ExcelPivotTableNaming.validateName(name);
    dev.erst.gridgrind.core.vocabulary.ExcelSheetNames.requireValid(sheetName, "sheetName");

    Optional<PivotHandle> handle = pivotByName(workbook, validatedName);
    if (handle.isEmpty() || !handle.orElseThrow().sheetName().equals(sheetName)) {
      throw new IllegalArgumentException(
          "pivot table not found on expected sheet: " + validatedName + "@" + sheetName);
    }
    deletePivotHandle(workbook, handle.orElseThrow(), allPivotTables(workbook), poiRelationRemover);
  }

  /** Returns factual pivot-table metadata selected by workbook-global name or all pivots. */
  public List<PivotTableReport> pivotTables(XssfWorkbook workbook, PivotTableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<PivotHandle> handles = selectHandles(workbook, selection);
    List<PivotTableReport> snapshots = new ArrayList<>(handles.size());
    for (PivotHandle handle : handles) {
      snapshots.add(snapshot(workbook.xssfWorkbook(), handle));
    }
    return List.copyOf(snapshots);
  }

  /** Returns integrity findings for the selected pivot-table set. */
  public List<AnalysisFindingReport> pivotTableHealthFindings(
      XssfWorkbook workbook, PivotTableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<PivotHandle> handles = selectHandles(workbook, selection);
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (PivotHandle handle : handles) {
      findings.addAll(pivotTableHealthFindings(workbook.xssfWorkbook(), handle));
    }
    findings.addAll(duplicateNameFindings(handles));
    return List.copyOf(new ArrayList<>(new LinkedHashSet<>(findings)));
  }

  public XSSFPivotTable createPivotTable(
      XssfWorkbook workbook,
      PivotTableSpec definition,
      ResolvedAuthoringSource source,
      CellReference anchor) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");

    var sheet = requiredSheet(workbook, definition.sheetName().value());
    sourceColumns(source.sheet(), source.area(), source.description());
    return switch (source) {
      case ResolvedAuthoringSource.Range range ->
          sheet.createPivotTable(range.area(), anchor, range.sheet());
      case ResolvedAuthoringSource.NamedRange namedRange ->
          sheet.createPivotTable(namedRange.namedRange(), anchor, namedRange.sheet());
      case ResolvedAuthoringSource.Table table -> sheet.createPivotTable(table.table(), anchor);
    };
  }

  public void normalizeCacheId(XSSFWorkbook workbook, XSSFPivotTable pivotTable) {
    XSSFPivotCache pivotCache = pivotTable.getPivotCache();
    if (pivotCache == null) {
      return;
    }
    Optional<XSSFPivotCacheDefinition> cacheDefinition = cacheDefinition(pivotTable);
    String currentRelationId = cacheDefinition.map(workbook::getRelationId).orElse(null);
    long currentId = pivotCache.getCTPivotCache().getCacheId();
    long maxOtherId = 0L;
    boolean duplicate = false;
    for (CTPivotCache cache : workbookPivotCaches(workbook)) {
      if (currentRelationId != null && currentRelationId.equals(cache.getId())) {
        continue;
      }
      maxOtherId = Math.max(maxOtherId, cache.getCacheId());
      if (cache.getCacheId() == currentId) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      pivotTable.getCTPivotTableDefinition().setCacheId(currentId);
      return;
    }
    long normalizedId = maxOtherId + 1L;
    pivotCache.getCTPivotCache().setCacheId(normalizedId);
    pivotTable.getCTPivotTableDefinition().setCacheId(normalizedId);
  }

  private List<PivotHandle> selectHandles(XssfWorkbook workbook, PivotTableSelector selection) {
    List<PivotHandle> all = allPivotTables(workbook);
    return switch (selection) {
      case PivotTableSelector.All ignored -> all;
      case PivotTableSelector.ByNames byNames -> selectHandlesByName(all, byNames.names());
      case PivotTableSelector.ByName byName -> selectHandlesByName(all, List.of(byName.name()));
      case PivotTableSelector.ByNameOnSheet scoped ->
          selectHandlesByName(all, List.of(scoped.name())).stream()
              .filter(handle -> handle.sheetName().equalsIgnoreCase(scoped.sheetName().value()))
              .toList();
    };
  }

  public List<PivotHandle> selectHandlesByName(List<PivotHandle> handles, List<String> names) {
    List<PivotHandle> selected = new ArrayList<>();
    for (String name : names) {
      findPivotHandleByName(handles, name).ifPresent(selected::add);
    }
    return List.copyOf(selected);
  }

  private Optional<PivotHandle> pivotByName(XssfWorkbook workbook, String name) {
    return findPivotHandleByName(allPivotTables(workbook), name);
  }

  private Optional<PivotHandle> findPivotHandleByName(List<PivotHandle> handles, String name) {
    String expected = ExcelPivotTableNaming.validateName(name).toUpperCase(Locale.ROOT);
    for (PivotHandle handle : handles) {
      if (resolvedName(handle).toUpperCase(Locale.ROOT).equals(expected)) {
        return Optional.of(handle);
      }
    }
    return Optional.empty();
  }

  public List<PivotHandle> allPivotTables(XssfWorkbook workbook) {
    List<PivotHandle> handles = new ArrayList<>();
    for (int sheetIndex = 0;
        sheetIndex < workbook.xssfWorkbook().getNumberOfSheets();
        sheetIndex++) {
      var sheet = workbook.xssfWorkbook().getSheetAt(sheetIndex);
      int ordinalOnSheet = 0;
      for (POIXMLDocumentPart relation : sheet.getRelations()) {
        if (!(relation instanceof XSSFPivotTable pivotTable)) {
          continue;
        }
        handles.add(
            new PivotHandle(sheetIndex, ordinalOnSheet, sheet.getSheetName(), sheet, pivotTable));
        ordinalOnSheet++;
      }
    }
    return List.copyOf(handles);
  }

  public void deletePivotHandle(XssfWorkbook workbook, PivotHandle handle) {
    deletePivotHandle(workbook, handle, allPivotTables(workbook), poiRelationRemover);
  }

  public boolean cacheDefinitionShared(
      XssfWorkbook workbook, PivotHandle current, XSSFPivotCacheDefinition cacheDefinition) {
    return cacheDefinitionShared(workbook, current, cacheDefinition, allPivotTables(workbook));
  }

  public boolean removePoiRelation(POIXMLDocumentPart parent, POIXMLDocumentPart child) {
    return removePoiRelation(parent, child, poiRelationRemover);
  }

  private static final String SYNTHETIC_PREFIX = "_GG_PIVOT_";

  public static Optional<PivotLocation> safeLocation(PivotHandle handle) {
    Optional<String> rawRange = rawLocationRange(handle);
    if (rawRange.isEmpty()) {
      return Optional.empty();
    }
    try {
      AreaReference area = new AreaReference(rawRange.orElseThrow(), SpreadsheetVersion.EXCEL2007);
      String locationRange = normalizeArea(area);
      return Optional.of(new PivotLocation(area.getFirstCell().formatAsString(), locationRange));
    } catch (RuntimeException exception) {
      return Optional.empty();
    }
  }

  public static Optional<String> rawLocationRange(PivotHandle handle) {
    var location = handle.table().getCTPivotTableDefinition().getLocation();
    return location == null ? Optional.empty() : blankAsOptional(location.getRef());
  }

  public static String resolvedName(PivotHandle handle) {
    return actualName(handle)
        .orElseGet(
            () ->
                syntheticName(
                    handle.sheetName(),
                    safeLocation(handle)
                        .map(PivotLocation::topLeftAddress)
                        .orElse("PIVOT_" + (handle.ordinalOnSheet() + 1))));
  }

  public static Optional<String> actualName(PivotHandle handle) {
    return blankAsOptional(handle.table().getCTPivotTableDefinition().getName());
  }

  private static String syntheticName(String sheetName, String topLeftAddress) {
    return SYNTHETIC_PREFIX + sanitize(sheetName) + '_' + sanitize(topLeftAddress);
  }

  public static String sanitize(String value) {
    StringBuilder builder = new StringBuilder(value.length());
    for (int index = 0; index < value.length(); index++) {
      char character = value.charAt(index);
      if (Character.isLetterOrDigit(character) || character == '_') {
        builder.append(character);
      } else {
        builder.append('_');
      }
    }
    return builder.toString();
  }

  public static AreaReference contiguousArea(String rawRange, String fieldName) {
    try {
      return new AreaReference(rawRange, SpreadsheetVersion.EXCEL2007);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          fieldName + " must be a contiguous A1-style range", exception);
    }
  }

  public static String normalizeArea(AreaReference area) {
    CellReference first =
        new CellReference(area.getFirstCell().getRow(), area.getFirstCell().getCol(), false, false);
    CellReference last =
        new CellReference(area.getLastCell().getRow(), area.getLastCell().getCol(), false, false);
    boolean singleCell = first.getRow() == last.getRow() && first.getCol() == last.getCol();
    return singleCell
        ? first.formatAsString()
        : first.formatAsString() + ":" + last.formatAsString();
  }

  public static String requireNonBlank(String value, String message) {
    if (value == null || value.isBlank()) {
      throw new IllegalArgumentException(message);
    }
    return value;
  }

  public static String nonBlankOrDefault(String value, String defaultValue) {
    return value == null || value.isBlank() ? defaultValue : value;
  }

  private static String normalizedResolvedName(PivotHandle handle) {
    return resolvedName(handle).toUpperCase(Locale.ROOT);
  }

  private static Optional<String> blankAsOptional(String value) {
    return value == null || value.isBlank() ? Optional.empty() : Optional.of(value);
  }

  private static List<AnalysisFindingReport> duplicateNameFindings(List<PivotHandle> handles) {
    Set<String> seenNames = new LinkedHashSet<>();
    Set<String> duplicateNames = new LinkedHashSet<>();
    for (PivotHandle handle : handles) {
      String normalizedName = normalizedResolvedName(handle);
      if (!seenNames.add(normalizedName)) {
        duplicateNames.add(normalizedName);
      }
    }
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (PivotHandle handle : handles) {
      if (!duplicateNames.contains(normalizedResolvedName(handle))) {
        continue;
      }
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.PIVOT_TABLE_DUPLICATE_NAME,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
              handle,
              "Pivot table name is not unique",
              "Multiple pivot tables share the same case-insensitive name, so exact-name selection is ambiguous.",
              List.of(resolvedName(handle))));
    }
    return List.copyOf(findings);
  }

  public static List<AnalysisFindingReport> pivotTableHealthFindings(
      XSSFWorkbook workbook, PivotHandle handle) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    if (actualName(handle).isEmpty()) {
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.WARNING,
              handle,
              "Pivot table name is missing",
              "The pivot table does not persist a name, so GridGrind assigned a synthetic identifier for readback.",
              List.of(resolvedName(handle))));
    }

    Optional<PivotLocation> location = safeLocation(handle);
    if (location.isEmpty()) {
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
              handle,
              "Pivot table location is malformed",
              "The pivot table location range could not be parsed.",
              rawLocationRange(handle)
                  .map(List::of)
                  .orElse(List.of("PIVOT_" + (handle.ordinalOnSheet() + 1)))));
      return List.copyOf(findings);
    }

    Optional<XSSFPivotCacheDefinition> cacheDefinition = cacheDefinition(handle.table());
    if (cacheDefinition.isEmpty()) {
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                  .PIVOT_TABLE_MISSING_CACHE_DEFINITION,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
              handle,
              "Pivot table is missing its cache definition relation",
              "The pivot table part no longer points at a pivot cache definition.",
              List.of(location.orElseThrow().locationRange())));
      return List.copyOf(findings);
    }

    if (cacheRecords(cacheDefinition.orElseThrow()).isEmpty()) {
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                  .PIVOT_TABLE_MISSING_CACHE_RECORDS,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
              handle,
              "Pivot table is missing its cache records relation",
              "The pivot cache definition does not point at pivot cache records.",
              List.of(location.orElseThrow().locationRange())));
    }

    CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
    if (workbookPivotCache(workbook, definition.getCacheId()).isEmpty()) {
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode
                  .PIVOT_TABLE_MISSING_WORKBOOK_CACHE,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
              handle,
              "Pivot table cache is not registered in workbook metadata",
              "The pivot table cacheId is missing from workbook.xml pivotCaches.",
              List.of(Long.toString(definition.getCacheId()))));
    }

    PivotTableReport snapshot = snapshot(workbook, handle);
    if (snapshot instanceof PivotTableReport.Unsupported unsupported) {
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.WARNING,
              handle,
              "Pivot table contains unsupported detail",
              unsupported.detail(),
              List.of(unsupported.anchor().locationRange().value())));
    }

    try {
      snapshotSource(workbook, handle.table());
    } catch (RuntimeException exception) {
      findings.add(
          finding(
              dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE,
              dev.erst.gridgrind.core.vocabulary.AnalysisSeverity.ERROR,
              handle,
              "Pivot table source is broken",
              Objects.requireNonNullElse(
                  exception.getMessage(), "The pivot source no longer resolves cleanly."),
              List.of(location.orElseThrow().locationRange())));
    }
    return List.copyOf(new ArrayList<>(new LinkedHashSet<>(findings)));
  }

  public static AnalysisFindingReport finding(
      dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode code,
      dev.erst.gridgrind.core.vocabulary.AnalysisSeverity severity,
      PivotHandle handle,
      String title,
      String message,
      List<String> evidence) {
    PivotLocation location = safeLocation(handle).orElse(null);
    AnalysisLocationReport analysisLocation =
        location == null
            ? new AnalysisLocationReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(handle.sheetName()))
            : new AnalysisLocationReport.Range(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(handle.sheetName()),
                location.locationRange());
    return new AnalysisFindingReport(
        code, severity, title, message, analysisLocation, List.copyOf(evidence));
  }

  public static void deletePivotHandle(
      XssfWorkbook workbook,
      PivotHandle handle,
      List<PivotHandle> allPivotTables,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    Optional<XSSFPivotCacheDefinition> cacheDefinition = cacheDefinition(handle.table());
    PackagePartName pivotTablePartName =
        handle.table().getPackagePart() == null
            ? null
            : handle.table().getPackagePart().getPartName();
    PackagePartName cacheDefinitionPartName =
        cacheDefinition.isEmpty() || cacheDefinition.orElseThrow().getPackagePart() == null
            ? null
            : cacheDefinition.orElseThrow().getPackagePart().getPartName();
    String cacheRelationId =
        cacheDefinition.isEmpty()
            ? null
            : handle.sheet().getWorkbook().getRelationId(cacheDefinition.orElseThrow());
    boolean sharedCache =
        cacheDefinition.isPresent()
            && cacheDefinitionShared(
                workbook, handle, cacheDefinition.orElseThrow(), allPivotTables);

    if (!removePoiRelation(handle.sheet(), handle.table(), poiRelationRemover)) {
      throw new IllegalStateException(
          "Failed to remove pivot table relation for '" + resolvedName(handle) + "'");
    }
    workbook.xssfWorkbook().getPivotTables().remove(handle.table());
    cleanupPackagePartIfUnused(workbook.xssfWorkbook().getPackage(), pivotTablePartName);

    if (!sharedCache && cacheDefinition.isPresent()) {
      removeWorkbookPivotCacheRegistration(
          workbook.xssfWorkbook(),
          handle.table().getCTPivotTableDefinition().getCacheId(),
          cacheRelationId);
      if (!removePoiRelation(
          workbook.xssfWorkbook(), cacheDefinition.orElseThrow(), poiRelationRemover)) {
        throw new IllegalStateException(
            "Failed to remove pivot cache definition relation for '" + resolvedName(handle) + "'");
      }
      cleanupPackagePartIfUnused(workbook.xssfWorkbook().getPackage(), cacheDefinitionPartName);
    }
    rebuildPivotTableRegistry(workbook.xssfWorkbook());
  }

  public static void removeWorkbookPivotCacheRegistration(
      XSSFWorkbook workbook, long cacheId, @Nullable String relationId) {
    if (!workbook.getCTWorkbook().isSetPivotCaches()) {
      return;
    }
    var pivotCaches = workbook.getCTWorkbook().getPivotCaches();
    for (int index = 0; index < pivotCaches.sizeOfPivotCacheArray(); index++) {
      CTPivotCache cache = pivotCaches.getPivotCacheArray(index);
      boolean matchesId = cache.getCacheId() == cacheId;
      boolean matchesRelation =
          relationId != null
              && relationId.equals(java.util.Objects.requireNonNullElse(cache.getId(), ""));
      if (matchesId || matchesRelation) {
        pivotCaches.removePivotCache(index);
        if (pivotCaches.sizeOfPivotCacheArray() == 0) {
          workbook.getCTWorkbook().unsetPivotCaches();
        }
        return;
      }
    }
  }

  public static boolean cacheDefinitionShared(
      XssfWorkbook workbook,
      PivotHandle current,
      XSSFPivotCacheDefinition cacheDefinition,
      List<PivotHandle> allPivotTables) {
    if (cacheDefinition.getPackagePart() == null) {
      return false;
    }
    String expectedPartName = cacheDefinition.getPackagePart().getPartName().getName();
    String currentPivotPartName = current.table().getPackagePart().getPartName().getName();
    for (PivotHandle handle : allPivotTables) {
      if (handle.table().getPackagePart().getPartName().getName().equals(currentPivotPartName)) {
        continue;
      }
      Optional<XSSFPivotCacheDefinition> otherCacheDefinition = cacheDefinition(handle.table());
      if (otherCacheDefinition.isEmpty()) {
        continue;
      }
      if (otherCacheDefinition
          .orElseThrow()
          .getPackagePart()
          .getPartName()
          .getName()
          .equals(expectedPartName)) {
        return true;
      }
    }
    return false;
  }

  public static boolean removePoiRelation(
      POIXMLDocumentPart parent,
      POIXMLDocumentPart child,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    return poiRelationRemover.test(parent, child);
  }

  public static void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg, @Nullable PackagePartName partName) {
    if (partName == null) {
      return;
    }
    for (PackagePart part :
        PoiPackageInspection.packageParts(pkg, "Failed to inspect package relationships")) {
      if (part.isRelationshipPart()) {
        continue;
      }
      for (org.apache.poi.openxml4j.opc.PackageRelationship relationship :
          PoiPackageInspection.relationships(part, "Failed to inspect package relationships")) {
        if (relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
          continue;
        }
        if (partName
            .getURI()
            .equals(
                org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
                    part.getPartName().getURI(), relationship.getTargetURI()))) {
          return;
        }
      }
    }
    if (pkg.containPart(partName)) {
      pkg.deletePartRecursive(partName);
    }
  }

  public static void primePivotTableAllocator(
      XSSFWorkbook workbook, Optional<XSSFPivotTable> allocationSentinel) {
    int highWaterMark = pivotTableIdHighWaterMark(workbook);
    List<XSSFPivotTable> registry = workbook.getPivotTables();
    if (registry.size() >= highWaterMark) {
      return;
    }
    if (registry.isEmpty()) {
      if (allocationSentinel.isEmpty()) {
        throw new IllegalStateException(
            "Pivot table allocation cannot advance because the workbook still contains pivot package numbering without any live pivot relations.");
      }
      while (registry.size() < highWaterMark) {
        registry.add(allocationSentinel.orElseThrow());
      }
      return;
    }
    XSSFPivotTable sentinel = registry.getLast();
    while (registry.size() < highWaterMark) {
      registry.add(sentinel);
    }
  }

  private static void rebuildPivotTableRegistry(XSSFWorkbook workbook) {
    List<XSSFPivotTable> registry = workbook.getPivotTables();
    registry.clear();
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      for (POIXMLDocumentPart relation : workbook.getSheetAt(sheetIndex).getRelations()) {
        if (relation instanceof XSSFPivotTable pivotTable) {
          registry.add(pivotTable);
        }
      }
    }
  }

  public static int pivotTableIdHighWaterMark(XSSFWorkbook workbook) {
    int maximum = 0;
    for (CTPivotCache cache : workbookPivotCaches(workbook)) {
      maximum = Math.max(maximum, Math.toIntExact(cache.getCacheId()));
    }
    for (PackagePart packagePart :
        PoiPackageInspection.packageParts(
            workbook.getPackage(), "Failed to inspect workbook package parts")) {
      maximum = Math.max(maximum, packagePartIndex(packagePart, "/xl/pivotTables/pivotTable"));
      maximum =
          Math.max(maximum, packagePartIndex(packagePart, "/xl/pivotCache/pivotCacheDefinition"));
      maximum =
          Math.max(maximum, packagePartIndex(packagePart, "/xl/pivotCache/pivotCacheRecords"));
    }
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      for (POIXMLDocumentPart relation : workbook.getSheetAt(sheetIndex).getRelations()) {
        if (!(relation instanceof XSSFPivotTable pivotTable)) {
          continue;
        }
        maximum = Math.max(maximum, packagePartIndex(pivotTable, "/xl/pivotTables/pivotTable"));
        Optional<XSSFPivotCacheDefinition> cacheDefinition = cacheDefinition(pivotTable);
        if (cacheDefinition.isPresent()) {
          maximum =
              Math.max(
                  maximum,
                  packagePartIndex(
                      cacheDefinition.orElseThrow(), "/xl/pivotCache/pivotCacheDefinition"));
          Optional<XSSFPivotCacheRecords> cacheRecords =
              cacheRecords(cacheDefinition.orElseThrow());
          if (cacheRecords.isPresent()) {
            maximum =
                Math.max(
                    maximum,
                    packagePartIndex(
                        cacheRecords.orElseThrow(), "/xl/pivotCache/pivotCacheRecords"));
          }
        }
      }
    }
    return maximum;
  }

  public static int packagePartIndex(POIXMLDocumentPart part, String prefix) {
    if (part.getPackagePart() == null) {
      return 0;
    }
    return packagePartIndex(part.getPackagePart(), prefix);
  }

  public static int packagePartIndex(PackagePart part, String prefix) {
    String name = part.getPartName().getName();
    if (!name.startsWith(prefix) || !name.endsWith(".xml")) {
      return 0;
    }
    String numericSuffix = name.substring(prefix.length(), name.length() - ".xml".length());
    try {
      return Integer.parseInt(numericSuffix);
    } catch (NumberFormatException exception) {
      return 0;
    }
  }

  public static PivotTableReport snapshot(XSSFWorkbook workbook, PivotHandle handle) {
    String name = resolvedName(handle);
    PivotLocation location = safeLocation(handle).orElse(null);
    PivotTableReport.Anchor anchor =
        location == null
            ? new PivotTableReport.Anchor(
                new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1"))
            : new PivotTableReport.Anchor(
                new dev.erst.gridgrind.core.vocabulary.CellAddress(location.topLeftAddress()),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel(location.locationRange()));
    if (location == null) {
      return unsupportedSnapshot(
          handle, name, anchor, "Pivot table location range is missing or malformed.");
    }
    try {
      CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
      List<String> sourceColumnNames = cacheFieldNames(handle.table());
      boolean valuesAxisOnColumns = false;
      List<PivotTableReport.Field> columnLabels = new ArrayList<>();
      if (definition.getColFields() != null) {
        for (CTField field : definition.getColFields().getFieldArray()) {
          if (field.getX() == -2) {
            valuesAxisOnColumns = true;
            continue;
          }
          columnLabels.add(sourceField(sourceColumnNames, field.getX()));
        }
      }
      List<PivotTableReport.DataField> dataFields =
          snapshotDataFields(workbook, definition, sourceColumnNames);
      if (dataFields.isEmpty()) {
        return unsupportedSnapshot(
            handle, name, anchor, "Pivot table does not contain any data fields.");
      }
      List<PivotTableReport.Field> rowFields =
          definition.getRowFields() == null
              ? List.of()
              : snapshotFields(
                  Objects.requireNonNull(
                      definition.getRowFields().getFieldArray(),
                      "pivot row field array must not be null"),
                  sourceColumnNames);
      List<PivotTableReport.Field> pageFields =
          definition.getPageFields() == null
              ? List.of()
              : snapshotPageFields(
                  Objects.requireNonNull(
                      definition.getPageFields().getPageFieldArray(),
                      "pivot page field array must not be null"),
                  sourceColumnNames);
      return new PivotTableReport.Supported(
          name,
          new dev.erst.gridgrind.core.vocabulary.SheetName(handle.sheetName()),
          anchor,
          snapshotSource(workbook, handle.table()),
          rowFields,
          columnLabels,
          pageFields,
          dataFields,
          valuesAxisOnColumns);
    } catch (IllegalArgumentException exception) {
      return unsupportedSnapshot(
          handle,
          name,
          anchor,
          Objects.requireNonNullElse(
              exception.getMessage(),
              "The pivot table could not be normalized into the modeled contract."));
    }
  }

  public static List<PivotTableReport.Field> snapshotFields(
      @Nullable CTField[] fields, List<String> sourceColumnNames) {
    if (fields == null || fields.length == 0) {
      return List.of();
    }
    List<PivotTableReport.Field> snapshots = new ArrayList<>(fields.length);
    for (CTField field : fields) {
      Objects.requireNonNull(field, "pivot field entries must not be null");
      snapshots.add(sourceField(sourceColumnNames, field.getX()));
    }
    return List.copyOf(snapshots);
  }

  public static List<PivotTableReport.Field> snapshotPageFields(
      @Nullable CTPageField[] pageFields, List<String> sourceColumnNames) {
    if (pageFields == null || pageFields.length == 0) {
      return List.of();
    }
    List<PivotTableReport.Field> snapshots = new ArrayList<>(pageFields.length);
    for (CTPageField pageField : pageFields) {
      Objects.requireNonNull(pageField, "pivot page-field entries must not be null");
      snapshots.add(sourceField(sourceColumnNames, pageField.getFld()));
    }
    return List.copyOf(snapshots);
  }

  public static List<PivotTableReport.DataField> snapshotDataFields(
      XSSFWorkbook workbook, CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    if (definition.getDataFields() == null
        || definition.getDataFields().sizeOfDataFieldArray() == 0) {
      return List.of();
    }
    List<PivotTableReport.DataField> dataFields = new ArrayList<>();
    for (CTDataField dataField : definition.getDataFields().getDataFieldArray()) {
      int sourceColumnIndex = Math.toIntExact(dataField.getFld());
      String sourceColumnName = sourceColumnName(sourceColumnNames, sourceColumnIndex);
      dataFields.add(
          new PivotTableReport.DataField(
              sourceColumnIndex,
              sourceColumnName,
              fromSubtotal(
                  dataField.getSubtotal() == null
                      ? DataConsolidateFunction.SUM.getValue()
                      : dataField.getSubtotal().intValue()),
              nonBlankOrDefault(dataField.getName(), sourceColumnName),
              numberFormat(workbook, dataField.isSetNumFmtId() ? dataField.getNumFmtId() : null)));
    }
    return List.copyOf(dataFields);
  }

  public static PivotTableReport.Unsupported unsupportedSnapshot(
      PivotHandle handle, String name, PivotTableReport.Anchor anchor, String detail) {
    return new PivotTableReport.Unsupported(
        name, new dev.erst.gridgrind.core.vocabulary.SheetName(handle.sheetName()), anchor, detail);
  }

  public static PivotTableReport.PivotSourceReport snapshotSource(
      XSSFWorkbook workbook, XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = requiredCacheDefinition(pivotTable);
    if (cacheDefinition.getCTPivotCacheDefinition().getCacheSource() == null
        || !cacheDefinition.getCTPivotCacheDefinition().getCacheSource().isSetWorksheetSource()) {
      throw new IllegalArgumentException("Pivot cache source is missing its worksheetSource.");
    }
    CTWorksheetSource worksheetSource =
        cacheDefinition.getCTPivotCacheDefinition().getCacheSource().getWorksheetSource();
    String sheetName =
        requireNonBlank(worksheetSource.getSheet(), "Pivot source sheet is missing.");
    if (worksheetSource.getRef() != null && !worksheetSource.getRef().isBlank()) {
      AreaReference area =
          new AreaReference(worksheetSource.getRef(), SpreadsheetVersion.EXCEL2007);
      return new PivotTableReport.PivotSourceReport.Range(
          new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(normalizeArea(area)));
    }

    String sourceName = requireNonBlank(worksheetSource.getName(), "Pivot source name is missing.");
    List<Name> namedRanges = matchingNamedRanges(workbook, sourceName, sheetName);
    if (namedRanges.size() == 1) {
      AreaReference area = namedRangeArea(namedRanges.getFirst());
      return new PivotTableReport.PivotSourceReport.NamedRange(
          new dev.erst.gridgrind.core.vocabulary.DefinedName(sourceName),
          new dev.erst.gridgrind.core.vocabulary.SheetName(
              sourceSheetName(area, namedRanges.getFirst(), sheetName)),
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(normalizeArea(area)));
    }
    if (namedRanges.size() > 1) {
      throw new IllegalArgumentException(
          "Pivot source name '"
              + sourceName
              + "' is ambiguous because multiple matching named ranges exist.");
    }

    Optional<XSSFTable> table = tableByName(workbook, sourceName, sheetName);
    if (table.isPresent()) {
      return new PivotTableReport.PivotSourceReport.Table(
          new dev.erst.gridgrind.core.vocabulary.TableName(sourceName),
          new dev.erst.gridgrind.core.vocabulary.SheetName(table.orElseThrow().getSheetName()),
          new dev.erst.gridgrind.core.vocabulary.RangeLabel(
              normalizeArea(
                  new AreaReference(
                      table.orElseThrow().getStartCellReference(),
                      table.orElseThrow().getEndCellReference(),
                      SpreadsheetVersion.EXCEL2007))));
    }
    throw new IllegalArgumentException(
        "Pivot source name '"
            + sourceName
            + "' does not resolve to an existing named range or table.");
  }

  public static List<String> cacheFieldNames(XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = requiredCacheDefinition(pivotTable);
    if (cacheDefinition.getCTPivotCacheDefinition().getCacheFields() == null) {
      throw new IllegalArgumentException("Pivot cache definition is missing cacheFields.");
    }
    List<String> fieldNames = new ArrayList<>();
    for (var cacheField :
        cacheDefinition.getCTPivotCacheDefinition().getCacheFields().getCacheFieldArray()) {
      fieldNames.add(requireNonBlank(cacheField.getName(), "Pivot cache field name is missing."));
    }
    return List.copyOf(fieldNames);
  }

  public static PivotTableReport.Field sourceField(
      List<String> sourceColumnNames, int sourceColumnIndex) {
    return new PivotTableReport.Field(
        sourceColumnIndex, sourceColumnName(sourceColumnNames, sourceColumnIndex));
  }

  public static ExcelPivotDataConsolidateFunction fromSubtotal(int subtotalValue) {
    for (DataConsolidateFunction function : DataConsolidateFunction.values()) {
      if (function.getValue() == subtotalValue) {
        return fromPoi(function);
      }
    }
    throw new IllegalArgumentException(
        "Unsupported pivot data consolidate function value: " + subtotalValue);
  }

  public static Optional<String> numberFormat(XSSFWorkbook workbook, @Nullable Long numFmtId) {
    if (numFmtId == null) {
      return Optional.empty();
    }
    short formatIndex = numFmtId > Short.MAX_VALUE ? Short.MAX_VALUE : (short) numFmtId.longValue();
    String format = workbook.createDataFormat().getFormat(formatIndex);
    return format == null || format.isBlank() ? Optional.empty() : Optional.of(format);
  }

  public static XSSFPivotCacheDefinition requiredCacheDefinition(XSSFPivotTable pivotTable) {
    return cacheDefinition(pivotTable)
        .orElseThrow(
            () ->
                new IllegalArgumentException(
                    "Pivot table is missing its cache definition relation."));
  }

  public static Optional<XSSFPivotCacheDefinition> cacheDefinition(XSSFPivotTable pivotTable) {
    return Optional.ofNullable(pivotTable.getPivotCacheDefinition());
  }

  public static Optional<XSSFPivotCacheRecords> cacheRecords(
      XSSFPivotCacheDefinition cacheDefinition) {
    return firstRelation(cacheDefinition, XSSFPivotCacheRecords.class);
  }

  public static <T extends POIXMLDocumentPart> Optional<T> firstRelation(
      POIXMLDocumentPart parent, Class<T> relationType) {
    for (POIXMLDocumentPart relation : parent.getRelations()) {
      if (relationType.isInstance(relation)) {
        return Optional.of(relationType.cast(relation));
      }
    }
    return Optional.empty();
  }

  public static Optional<CTPivotCache> workbookPivotCache(XSSFWorkbook workbook, long cacheId) {
    for (CTPivotCache cache : workbookPivotCaches(workbook)) {
      if (cache.getCacheId() == cacheId) {
        return Optional.of(cache);
      }
    }
    return Optional.empty();
  }

  private static List<CTPivotCache> workbookPivotCaches(XSSFWorkbook workbook) {
    if (!workbook.getCTWorkbook().isSetPivotCaches()) {
      return List.of();
    }
    return List.of(workbook.getCTWorkbook().getPivotCaches().getPivotCacheArray());
  }

  private static ResolvedAuthoringSource resolveAuthoringSource(
      XssfWorkbook workbook, PivotTableSpec.PivotSourceSpec source) {
    return switch (source) {
      case PivotTableSpec.PivotSourceSpec.Range range -> {
        XSSFSheet sheet = requiredSheet(workbook, range.sheetName().value());
        AreaReference area = contiguousArea(range.range().value(), "range");
        yield ResolvedAuthoringSource.range(sheet, area);
      }
      case PivotTableSpec.PivotSourceSpec.NamedRange namedRange -> {
        List<Name> matches =
            matchingNamedRanges(workbook.xssfWorkbook(), namedRange.name().value(), null);
        if (matches.isEmpty()) {
          throw new IllegalArgumentException(
              "pivot source named range not found: " + namedRange.name().value());
        }
        if (matches.size() > 1) {
          throw new IllegalArgumentException(
              "pivot source named range is ambiguous: " + namedRange.name().value());
        }
        Name resolved = matches.getFirst();
        AreaReference area = namedRangeArea(resolved);
        XSSFSheet sheet = requiredSheet(workbook, sourceSheetName(area, resolved, null));
        yield ResolvedAuthoringSource.namedRange(sheet, area, resolved);
      }
      case PivotTableSpec.PivotSourceSpec.Table table -> {
        XSSFTable resolved = requiredTableByName(workbook, table.name().value());
        XSSFSheet sheet = requiredSheet(workbook, resolved.getSheetName());
        yield ResolvedAuthoringSource.table(
            sheet,
            new AreaReference(
                resolved.getStartCellReference(),
                resolved.getEndCellReference(),
                org.apache.poi.ss.SpreadsheetVersion.EXCEL2007),
            resolved);
      }
    };
  }

  public static XSSFTable requiredTableByName(XssfWorkbook workbook, String name) {
    return tableByName(workbook.xssfWorkbook(), name, null)
        .orElseThrow(() -> new IllegalArgumentException("pivot source table not found: " + name));
  }

  public static Optional<XSSFTable> tableByName(
      XSSFWorkbook workbook, String name, @Nullable String preferredSheetName) {
    String expected = name.toUpperCase(Locale.ROOT);
    XSSFTable match = null;
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
      for (XSSFTable table : sheet.getTables()) {
        if (!table.getName().toUpperCase(Locale.ROOT).equals(expected)) {
          continue;
        }
        if (preferredSheetName != null && !sheet.getSheetName().equals(preferredSheetName)) {
          continue;
        }
        if (match != null) {
          throw new IllegalArgumentException("pivot source table is ambiguous: " + name);
        }
        match = table;
      }
    }
    return Optional.ofNullable(match);
  }

  public static List<Name> matchingNamedRanges(
      XSSFWorkbook workbook, String name, @Nullable String sheetNameHint) {
    List<Name> matches = new ArrayList<>();
    for (Name candidate : workbook.getAllNames()) {
      if (!candidate.getNameName().equalsIgnoreCase(name)) {
        continue;
      }
      if (sheetNameHint != null && candidate.getSheetIndex() >= 0) {
        String candidateSheetName = workbook.getSheetName(candidate.getSheetIndex());
        if (!sheetNameHint.equals(candidateSheetName)) {
          continue;
        }
      }
      matches.add(candidate);
    }
    return List.copyOf(matches);
  }

  public static AreaReference namedRangeArea(Name namedRange) {
    String formula = java.util.Objects.requireNonNullElse(namedRange.getRefersToFormula(), "");
    if (formula.isBlank()) {
      throw new IllegalArgumentException(
          "pivot source named range " + namedRange.getNameName() + " does not define an area");
    }
    try {
      return new AreaReference(
          NamedRangeAdapter.normalizeAreaFormulaForPoi(formula),
          org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "pivot source named range "
              + namedRange.getNameName()
              + " does not define a rectangular area",
          exception);
    }
  }

  public static String sourceSheetName(
      AreaReference area, @Nullable Name namedRange, @Nullable String fallbackSheetName) {
    String areaSheetName = area.getFirstCell().getSheetName();
    if (areaSheetName != null) {
      return areaSheetName;
    }
    if (namedRange != null && namedRange.getSheetIndex() >= 0) {
      return namedRange.getSheetName();
    }
    if (fallbackSheetName != null && !fallbackSheetName.isBlank()) {
      return fallbackSheetName;
    }
    throw new IllegalArgumentException("pivot source area does not identify its sheet");
  }

  public static SourceColumns sourceColumns(
      XSSFSheet sheet, AreaReference area, String description) {
    CellReference firstCell = area.getFirstCell();
    CellReference lastCell = area.getLastCell();
    requireBodyRows(firstCell, lastCell, description);
    Row headerRow = requiredHeaderRow(sheet, firstCell, description);

    List<SourceColumn> columns = new ArrayList<>();
    Set<String> seenNames = new java.util.LinkedHashSet<>();
    for (int columnIndex = firstCell.getCol(); columnIndex <= lastCell.getCol(); columnIndex++) {
      String name = requiredHeaderName(headerRow, columnIndex, description);
      requireUniqueHeaderName(seenNames, name, description);
      columns.add(new SourceColumn(name, columnIndex - firstCell.getCol()));
    }
    return new SourceColumns(columns);
  }

  private static void requireBodyRows(
      CellReference firstCell, CellReference lastCell, String description) {
    if (lastCell.getRow() <= firstCell.getRow()) {
      throw new IllegalArgumentException(
          "pivot source " + description + " must include a header row plus at least one data row");
    }
  }

  private static Row requiredHeaderRow(
      XSSFSheet sheet, CellReference firstCell, String description) {
    Row headerRow = sheet.getRow(firstCell.getRow());
    if (headerRow == null) {
      throw new IllegalArgumentException(
          "pivot source " + description + " is missing its header row");
    }
    return headerRow;
  }

  private static String requiredHeaderName(Row headerRow, int columnIndex, String description) {
    var cell = headerRow.getCell(columnIndex);
    if (cell == null || cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) {
      throw new IllegalArgumentException(
          "pivot source " + description + " header cells must all be strings");
    }
    String name = cell.getStringCellValue();
    if (name.isBlank()) {
      throw new IllegalArgumentException(
          "pivot source " + description + " contains a blank header cell");
    }
    return name;
  }

  private static void requireUniqueHeaderName(
      Set<String> seenNames, String name, String description) {
    if (!seenNames.add(name.toUpperCase(Locale.ROOT))) {
      throw new IllegalArgumentException(
          "pivot source " + description + " header row must be unique case-insensitively: " + name);
    }
  }

  public static String sourceColumnName(List<String> sourceColumnNames, int sourceColumnIndex) {
    if (sourceColumnIndex < 0 || sourceColumnIndex >= sourceColumnNames.size()) {
      throw new IllegalArgumentException(
          "pivot source column index is out of bounds: " + sourceColumnIndex);
    }
    return sourceColumnNames.get(sourceColumnIndex);
  }

  public static XSSFSheet requiredSheet(XssfWorkbook workbook, String sheetName) {
    XSSFSheet sheet = workbook.xssfWorkbook().getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  public static DataConsolidateFunction toPoi(ExcelPivotDataConsolidateFunction function) {
    if (function == null) {
      throw new IllegalArgumentException("Unsupported GridGrind pivot data function: null");
    }
    return DataConsolidateFunction.valueOf(function.name());
  }

  public static ExcelPivotDataConsolidateFunction fromPoi(DataConsolidateFunction function) {
    if (function == null) {
      throw new IllegalArgumentException("Unsupported Apache POI pivot data function: null");
    }
    return ExcelPivotDataConsolidateFunction.valueOf(function.name());
  }

  public record PivotHandle(
      int sheetIndex,
      int ordinalOnSheet,
      String sheetName,
      XSSFSheet sheet,
      XSSFPivotTable table) {}

  public record PivotLocation(String topLeftAddress, String locationRange) {}

  public record SourceColumn(String name, int relativeIndex) {}

  public record SourceColumns(List<SourceColumn> columns) {
    public SourceColumns {
      columns = List.copyOf(columns);
    }

    public int relativeIndex(String name) {
      String expected = name.toUpperCase(Locale.ROOT);
      for (SourceColumn column : columns) {
        if (column.name().toUpperCase(Locale.ROOT).equals(expected)) {
          return column.relativeIndex();
        }
      }
      throw new IllegalArgumentException("pivot source column not found: " + name);
    }
  }

  public sealed interface ResolvedAuthoringSource
      permits ResolvedAuthoringSource.Range,
          ResolvedAuthoringSource.NamedRange,
          ResolvedAuthoringSource.Table {
    XSSFSheet sheet();

    AreaReference area();

    String description();

    static Range range(XSSFSheet sheet, AreaReference area) {
      return new Range(sheet, area, sheet.getSheetName() + "!" + area.formatAsString());
    }

    static NamedRange namedRange(XSSFSheet sheet, AreaReference area, Name namedRange) {
      return new NamedRange(sheet, area, "named range " + namedRange.getNameName(), namedRange);
    }

    static Table table(XSSFSheet sheet, AreaReference area, XSSFTable table) {
      return new Table(sheet, area, "table " + table.getName(), table);
    }

    record Range(XSSFSheet sheet, AreaReference area, String description)
        implements ResolvedAuthoringSource {
      public Range {
        Objects.requireNonNull(sheet, "sheet must not be null");
        Objects.requireNonNull(area, "area must not be null");
        Objects.requireNonNull(description, "description must not be null");
      }
    }

    record NamedRange(XSSFSheet sheet, AreaReference area, String description, Name namedRange)
        implements ResolvedAuthoringSource {
      public NamedRange {
        Objects.requireNonNull(sheet, "sheet must not be null");
        Objects.requireNonNull(area, "area must not be null");
        Objects.requireNonNull(description, "description must not be null");
        Objects.requireNonNull(namedRange, "namedRange must not be null");
      }
    }

    record Table(XSSFSheet sheet, AreaReference area, String description, XSSFTable table)
        implements ResolvedAuthoringSource {
      public Table {
        Objects.requireNonNull(sheet, "sheet must not be null");
        Objects.requireNonNull(area, "area must not be null");
        Objects.requireNonNull(description, "description must not be null");
        Objects.requireNonNull(table, "table must not be null");
      }
    }
  }
}
