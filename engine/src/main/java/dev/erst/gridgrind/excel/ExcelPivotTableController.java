package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;

/** Reads, writes, and analyzes workbook pivot tables within the POI-supported XSSF surface. */
final class ExcelPivotTableController {
  private static final String SYNTHETIC_PREFIX = "_GG_PIVOT_";
  private final BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover;

  ExcelPivotTableController() {
    this(PoiRelationRemoval.defaultRemover());
  }

  ExcelPivotTableController(
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    this.poiRelationRemover =
        Objects.requireNonNull(poiRelationRemover, "poiRelationRemover must not be null");
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  void setPivotTable(ExcelWorkbook workbook, ExcelPivotTableDefinition definition) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    PivotHandle existing = pivotByName(workbook, definition.name());
    if (existing != null && !existing.sheetName().equals(definition.sheetName())) {
      throw new IllegalArgumentException(
          "pivot table name already exists on a different sheet: " + definition.name());
    }

    ResolvedAuthoringSource source = resolveAuthoringSource(workbook, definition.source());
    SourceColumns columns = sourceColumns(source.sheet(), source.area(), source.description());
    CellReference anchor = new CellReference(definition.anchor().topLeftAddress());
    requireSupportedReportFilterAnchor(definition, anchor);

    if (existing != null) {
      deletePivotHandle(workbook, existing);
    }

    primePivotTableAllocator(workbook.xssfWorkbook(), existing == null ? null : existing.table());
    try {
      XSSFPivotTable pivotTable = createPivotTable(workbook, definition, source, anchor);
      normalizeCacheId(workbook.xssfWorkbook(), pivotTable);
      pivotTable.getCTPivotTableDefinition().setName(definition.name());
      applyPivotFields(pivotTable, definition, columns);
    } finally {
      rebuildPivotTableRegistry(workbook.xssfWorkbook());
    }
  }

  private void requireSupportedReportFilterAnchor(
      ExcelPivotTableDefinition definition, CellReference anchor) {
    if (!definition.reportFilters().isEmpty() && anchor.getRow() < 2) {
      throw new IllegalArgumentException(
          "pivot tables with reportFilters require anchor.topLeftAddress on row 3 or lower because Apache POI moves page filters below the first two rows");
    }
  }

  private void applyPivotFields(
      XSSFPivotTable pivotTable, ExcelPivotTableDefinition definition, SourceColumns columns) {
    for (String rowLabel : definition.rowLabels()) {
      pivotTable.addRowLabel(columns.relativeIndex(rowLabel));
    }
    for (String columnLabel : definition.columnLabels()) {
      pivotTable.addColLabel(columns.relativeIndex(columnLabel));
    }
    for (String reportFilter : definition.reportFilters()) {
      pivotTable.addReportFilter(columns.relativeIndex(reportFilter));
    }
    for (ExcelPivotTableDefinition.DataField dataField : definition.dataFields()) {
      pivotTable.addColumnLabel(
          dataField.function().toPoiFunction(),
          columns.relativeIndex(dataField.sourceColumnName()),
          dataField.displayName(),
          dataField.valueFormat());
    }
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet name. */
  void deletePivotTable(ExcelWorkbook workbook, String name, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String validatedName = ExcelPivotTableNaming.validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");

    PivotHandle handle = pivotByName(workbook, validatedName);
    if (handle == null || !handle.sheetName().equals(sheetName)) {
      throw new IllegalArgumentException(
          "pivot table not found on expected sheet: " + validatedName + "@" + sheetName);
    }
    deletePivotHandle(workbook, handle);
  }

  /** Returns factual pivot-table metadata selected by workbook-global name or all pivots. */
  List<ExcelPivotTableSnapshot> pivotTables(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<PivotHandle> handles = selectHandles(workbook, selection);
    List<ExcelPivotTableSnapshot> snapshots = new ArrayList<>(handles.size());
    for (PivotHandle handle : handles) {
      snapshots.add(snapshot(workbook.xssfWorkbook(), handle));
    }
    return List.copyOf(snapshots);
  }

  /** Returns integrity findings for the selected pivot-table set. */
  List<WorkbookAnalysis.AnalysisFinding> pivotTableHealthFindings(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<PivotHandle> handles = selectHandles(workbook, selection);
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (PivotHandle handle : handles) {
      findings.addAll(pivotTableHealthFindings(workbook.xssfWorkbook(), handle));
    }
    findings.addAll(duplicateNameFindings(handles));
    return List.copyOf(new ArrayList<>(new LinkedHashSet<>(findings)));
  }

  private List<WorkbookAnalysis.AnalysisFinding> duplicateNameFindings(List<PivotHandle> handles) {
    Set<String> seenNames = new LinkedHashSet<>();
    Set<String> duplicateNames = new LinkedHashSet<>();
    for (PivotHandle handle : handles) {
      String normalizedName = normalizedResolvedName(handle);
      if (!seenNames.add(normalizedName)) {
        duplicateNames.add(normalizedName);
      }
    }
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (PivotHandle handle : handles) {
      if (!duplicateNames.contains(normalizedResolvedName(handle))) {
        continue;
      }
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_DUPLICATE_NAME,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              handle,
              "Pivot table name is not unique",
              "Multiple pivot tables share the same case-insensitive name, so exact-name selection is ambiguous.",
              List.of(resolvedName(handle))));
    }
    return List.copyOf(findings);
  }

  private List<WorkbookAnalysis.AnalysisFinding> pivotTableHealthFindings(
      XSSFWorkbook workbook, PivotHandle handle) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    if (actualName(handle) == null) {
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
              WorkbookAnalysis.AnalysisSeverity.WARNING,
              handle,
              "Pivot table name is missing",
              "The pivot table does not persist a name, so GridGrind assigned a synthetic identifier for readback.",
              List.of(resolvedName(handle))));
    }

    PivotLocation location = safeLocation(handle);
    if (location == null) {
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              handle,
              "Pivot table location is malformed",
              "The pivot table location range could not be parsed.",
              List.of(rawLocationRange(handle))));
      return List.copyOf(findings);
    }

    XSSFPivotCacheDefinition cacheDefinition = cacheDefinition(handle.table());
    if (cacheDefinition == null) {
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_DEFINITION,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              handle,
              "Pivot table is missing its cache definition relation",
              "The pivot table part no longer points at a pivot cache definition.",
              List.of(location.locationRange())));
      return List.copyOf(findings);
    }

    if (cacheRecords(cacheDefinition) == null) {
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_RECORDS,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              handle,
              "Pivot table is missing its cache records relation",
              "The pivot cache definition does not point at pivot cache records.",
              List.of(location.locationRange())));
    }

    CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
    if (workbookPivotCache(workbook, definition.getCacheId()) == null) {
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_WORKBOOK_CACHE,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              handle,
              "Pivot table cache is not registered in workbook metadata",
              "The pivot table cacheId is missing from workbook.xml pivotCaches.",
              List.of(Long.toString(definition.getCacheId()))));
    }

    ExcelPivotTableSnapshot snapshot = snapshot(workbook, handle);
    if (snapshot instanceof ExcelPivotTableSnapshot.Unsupported unsupported) {
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
              WorkbookAnalysis.AnalysisSeverity.WARNING,
              handle,
              "Pivot table contains unsupported detail",
              unsupported.detail(),
              List.of(unsupported.anchor().locationRange())));
    }

    try {
      snapshotSource(workbook, handle.table());
    } catch (RuntimeException exception) {
      findings.add(
          finding(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              handle,
              "Pivot table source is broken",
              Objects.requireNonNullElse(
                  exception.getMessage(), "The pivot source no longer resolves cleanly."),
              List.of(location.locationRange())));
    }
    return List.copyOf(new ArrayList<>(new LinkedHashSet<>(findings)));
  }

  private WorkbookAnalysis.AnalysisFinding finding(
      WorkbookAnalysis.AnalysisFindingCode code,
      WorkbookAnalysis.AnalysisSeverity severity,
      PivotHandle handle,
      String title,
      String message,
      List<String> evidence) {
    PivotLocation location = safeLocation(handle);
    WorkbookAnalysis.AnalysisLocation analysisLocation =
        location == null
            ? new WorkbookAnalysis.AnalysisLocation.Sheet(handle.sheetName())
            : new WorkbookAnalysis.AnalysisLocation.Range(
                handle.sheetName(), location.locationRange());
    return new WorkbookAnalysis.AnalysisFinding(
        code, severity, title, message, analysisLocation, List.copyOf(evidence));
  }

  private XSSFPivotTable createPivotTable(
      ExcelWorkbook workbook,
      ExcelPivotTableDefinition definition,
      ResolvedAuthoringSource source,
      CellReference anchor) {
    XSSFSheet sheet = requiredSheet(workbook, definition.sheetName());
    return switch (source.kind()) {
      case RANGE -> sheet.createPivotTable(source.area(), anchor, source.sheet());
      case NAMED_RANGE ->
          sheet.createPivotTable(
              Objects.requireNonNull(source.namedRange(), "namedRange must not be null"),
              anchor,
              source.sheet());
      case TABLE ->
          sheet.createPivotTable(
              Objects.requireNonNull(source.table(), "table must not be null"), anchor);
    };
  }

  private void normalizeCacheId(XSSFWorkbook workbook, XSSFPivotTable pivotTable) {
    XSSFPivotCache pivotCache = pivotTable.getPivotCache();
    if (pivotCache == null) {
      return;
    }
    XSSFPivotCacheDefinition cacheDefinition = cacheDefinition(pivotTable);
    String currentRelationId =
        cacheDefinition == null ? null : workbook.getRelationId(cacheDefinition);
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

  private void deletePivotHandle(ExcelWorkbook workbook, PivotHandle handle) {
    XSSFPivotCacheDefinition cacheDefinition = cacheDefinition(handle.table());
    org.apache.poi.openxml4j.opc.PackagePartName pivotTablePartName =
        handle.table().getPackagePart() == null
            ? null
            : handle.table().getPackagePart().getPartName();
    org.apache.poi.openxml4j.opc.PackagePartName cacheDefinitionPartName =
        cacheDefinition == null || cacheDefinition.getPackagePart() == null
            ? null
            : cacheDefinition.getPackagePart().getPartName();
    String cacheRelationId =
        cacheDefinition == null
            ? null
            : handle.sheet().getWorkbook().getRelationId(cacheDefinition);
    boolean sharedCache =
        cacheDefinition != null && cacheDefinitionShared(workbook, handle, cacheDefinition);

    if (!removePoiRelation(handle.sheet(), handle.table())) {
      throw new IllegalStateException(
          "Failed to remove pivot table relation for '" + resolvedName(handle) + "'");
    }
    workbook.xssfWorkbook().getPivotTables().remove(handle.table());
    cleanupPackagePartIfUnused(workbook.xssfWorkbook().getPackage(), pivotTablePartName);

    if (!sharedCache && cacheDefinition != null) {
      removeWorkbookPivotCacheRegistration(
          workbook.xssfWorkbook(),
          handle.table().getCTPivotTableDefinition().getCacheId(),
          cacheRelationId);
      if (!removePoiRelation(workbook.xssfWorkbook(), cacheDefinition)) {
        throw new IllegalStateException(
            "Failed to remove pivot cache definition relation for '" + resolvedName(handle) + "'");
      }
      cleanupPackagePartIfUnused(workbook.xssfWorkbook().getPackage(), cacheDefinitionPartName);
    }
    rebuildPivotTableRegistry(workbook.xssfWorkbook());
  }

  private void removeWorkbookPivotCacheRegistration(
      XSSFWorkbook workbook, long cacheId, String relationId) {
    if (!workbook.getCTWorkbook().isSetPivotCaches()) {
      return;
    }
    var pivotCaches = workbook.getCTWorkbook().getPivotCaches();
    for (int index = 0; index < pivotCaches.sizeOfPivotCacheArray(); index++) {
      CTPivotCache cache = pivotCaches.getPivotCacheArray(index);
      boolean matchesId = cache.getCacheId() == cacheId;
      boolean matchesRelation =
          relationId != null && relationId.equals(Objects.requireNonNullElse(cache.getId(), ""));
      if (matchesId || matchesRelation) {
        pivotCaches.removePivotCache(index);
        if (pivotCaches.sizeOfPivotCacheArray() == 0) {
          workbook.getCTWorkbook().unsetPivotCaches();
        }
        return;
      }
    }
  }

  private boolean cacheDefinitionShared(
      ExcelWorkbook workbook, PivotHandle current, XSSFPivotCacheDefinition cacheDefinition) {
    String expectedPartName = cacheDefinition.getPackagePart().getPartName().getName();
    String currentPivotPartName = current.table().getPackagePart().getPartName().getName();
    for (PivotHandle handle : allPivotTables(workbook)) {
      if (handle.table().getPackagePart().getPartName().getName().equals(currentPivotPartName)) {
        continue;
      }
      XSSFPivotCacheDefinition otherCacheDefinition = cacheDefinition(handle.table());
      if (otherCacheDefinition == null) {
        continue;
      }
      if (otherCacheDefinition.getPackagePart().getPartName().getName().equals(expectedPartName)) {
        return true;
      }
    }
    return false;
  }

  private boolean removePoiRelation(POIXMLDocumentPart parent, POIXMLDocumentPart child) {
    return poiRelationRemover.test(parent, child);
  }

  private void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg,
      org.apache.poi.openxml4j.opc.PackagePartName partName) {
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

  private void primePivotTableAllocator(XSSFWorkbook workbook, XSSFPivotTable allocationSentinel) {
    int highWaterMark = pivotTableIdHighWaterMark(workbook);
    List<XSSFPivotTable> registry = workbook.getPivotTables();
    if (registry.size() >= highWaterMark) {
      return;
    }
    if (registry.isEmpty()) {
      if (allocationSentinel == null) {
        throw new IllegalStateException(
            "Pivot table allocation cannot advance because the workbook still contains pivot package numbering without any live pivot relations.");
      }
      while (registry.size() < highWaterMark) {
        registry.add(allocationSentinel);
      }
      return;
    }
    XSSFPivotTable sentinel = registry.getLast();
    while (registry.size() < highWaterMark) {
      registry.add(sentinel);
    }
  }

  private void rebuildPivotTableRegistry(XSSFWorkbook workbook) {
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

  private int pivotTableIdHighWaterMark(XSSFWorkbook workbook) {
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
        XSSFPivotCacheDefinition cacheDefinition = cacheDefinition(pivotTable);
        if (cacheDefinition != null) {
          maximum =
              Math.max(
                  maximum,
                  packagePartIndex(cacheDefinition, "/xl/pivotCache/pivotCacheDefinition"));
          XSSFPivotCacheRecords cacheRecords = cacheRecords(cacheDefinition);
          if (cacheRecords != null) {
            maximum =
                Math.max(
                    maximum, packagePartIndex(cacheRecords, "/xl/pivotCache/pivotCacheRecords"));
          }
        }
      }
    }
    return maximum;
  }

  private int packagePartIndex(POIXMLDocumentPart part, String prefix) {
    if (part.getPackagePart() == null) {
      return 0;
    }
    return packagePartIndex(part.getPackagePart(), prefix);
  }

  private int packagePartIndex(PackagePart part, String prefix) {
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

  private List<PivotHandle> selectHandles(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    List<PivotHandle> all = allPivotTables(workbook);
    return switch (selection) {
      case ExcelPivotTableSelection.All ignored -> all;
      case ExcelPivotTableSelection.ByNames byNames -> selectHandlesByName(all, byNames.names());
    };
  }

  private List<PivotHandle> selectHandlesByName(List<PivotHandle> handles, List<String> names) {
    List<PivotHandle> selected = new ArrayList<>();
    for (String name : names) {
      PivotHandle handle = findPivotHandleByName(handles, name);
      if (handle != null) {
        selected.add(handle);
      }
    }
    return List.copyOf(selected);
  }

  private PivotHandle pivotByName(ExcelWorkbook workbook, String name) {
    return findPivotHandleByName(allPivotTables(workbook), name);
  }

  private PivotHandle findPivotHandleByName(List<PivotHandle> handles, String name) {
    String expected = ExcelPivotTableNaming.validateName(name).toUpperCase(Locale.ROOT);
    for (PivotHandle handle : handles) {
      if (resolvedName(handle).toUpperCase(Locale.ROOT).equals(expected)) {
        return handle;
      }
    }
    return null;
  }

  private List<PivotHandle> allPivotTables(ExcelWorkbook workbook) {
    List<PivotHandle> handles = new ArrayList<>();
    for (int sheetIndex = 0;
        sheetIndex < workbook.xssfWorkbook().getNumberOfSheets();
        sheetIndex++) {
      XSSFSheet sheet = workbook.xssfWorkbook().getSheetAt(sheetIndex);
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

  private ExcelPivotTableSnapshot snapshot(XSSFWorkbook workbook, PivotHandle handle) {
    String name = resolvedName(handle);
    PivotLocation location = safeLocation(handle);
    ExcelPivotTableSnapshot.Anchor anchor =
        location == null
            ? new ExcelPivotTableSnapshot.Anchor("A1", "A1")
            : new ExcelPivotTableSnapshot.Anchor(
                location.topLeftAddress(), location.locationRange());
    if (location == null) {
      return unsupportedSnapshot(
          handle, name, anchor, "Pivot table location range is missing or malformed.");
    }
    try {
      CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
      List<String> sourceColumnNames = cacheFieldNames(handle.table());
      ColumnAxisSnapshot columns = snapshotColumnLabels(definition, sourceColumnNames);
      List<ExcelPivotTableSnapshot.DataField> dataFields =
          snapshotDataFields(workbook, definition, sourceColumnNames);
      if (dataFields.isEmpty()) {
        return unsupportedSnapshot(
            handle, name, anchor, "Pivot table does not contain any data fields.");
      }
      return new ExcelPivotTableSnapshot.Supported(
          name,
          handle.sheetName(),
          anchor,
          snapshotSource(workbook, handle.table()),
          snapshotFields(
              definition.getRowFields() == null ? null : definition.getRowFields().getFieldArray(),
              sourceColumnNames),
          columns.columnLabels(),
          snapshotPageFields(
              definition.getPageFields() == null
                  ? null
                  : definition.getPageFields().getPageFieldArray(),
              sourceColumnNames),
          dataFields,
          columns.valuesAxisOnColumns());
    } catch (RuntimeException exception) {
      return unsupportedSnapshot(
          handle,
          name,
          anchor,
          Objects.requireNonNullElse(
              exception.getMessage(),
              "The pivot table could not be normalized into the modeled contract."));
    }
  }

  private ColumnAxisSnapshot snapshotColumnLabels(
      CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    boolean valuesAxisOnColumns = false;
    List<ExcelPivotTableSnapshot.Field> columnLabels = new ArrayList<>();
    if (definition.getColFields() != null) {
      for (CTField field : definition.getColFields().getFieldArray()) {
        if (field.getX() == -2) {
          valuesAxisOnColumns = true;
          continue;
        }
        columnLabels.add(sourceField(sourceColumnNames, field.getX()));
      }
    }
    return new ColumnAxisSnapshot(List.copyOf(columnLabels), valuesAxisOnColumns);
  }

  private List<ExcelPivotTableSnapshot.Field> snapshotFields(
      CTField[] fields, List<String> sourceColumnNames) {
    if (fields == null || fields.length == 0) {
      return List.of();
    }
    List<ExcelPivotTableSnapshot.Field> snapshots = new ArrayList<>(fields.length);
    for (CTField field : fields) {
      snapshots.add(sourceField(sourceColumnNames, field.getX()));
    }
    return List.copyOf(snapshots);
  }

  private List<ExcelPivotTableSnapshot.Field> snapshotPageFields(
      CTPageField[] pageFields, List<String> sourceColumnNames) {
    if (pageFields == null || pageFields.length == 0) {
      return List.of();
    }
    List<ExcelPivotTableSnapshot.Field> snapshots = new ArrayList<>(pageFields.length);
    for (CTPageField pageField : pageFields) {
      snapshots.add(sourceField(sourceColumnNames, pageField.getFld()));
    }
    return List.copyOf(snapshots);
  }

  private List<ExcelPivotTableSnapshot.DataField> snapshotDataFields(
      XSSFWorkbook workbook, CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    if (definition.getDataFields() == null
        || definition.getDataFields().sizeOfDataFieldArray() == 0) {
      return List.of();
    }
    List<ExcelPivotTableSnapshot.DataField> dataFields = new ArrayList<>();
    for (CTDataField dataField : definition.getDataFields().getDataFieldArray()) {
      int sourceColumnIndex = Math.toIntExact(dataField.getFld());
      String sourceColumnName = sourceColumnName(sourceColumnNames, sourceColumnIndex);
      dataFields.add(
          new ExcelPivotTableSnapshot.DataField(
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

  private ExcelPivotTableSnapshot.Unsupported unsupportedSnapshot(
      PivotHandle handle, String name, ExcelPivotTableSnapshot.Anchor anchor, String detail) {
    return new ExcelPivotTableSnapshot.Unsupported(name, handle.sheetName(), anchor, detail);
  }

  private ExcelPivotTableSnapshot.Source snapshotSource(
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
      return new ExcelPivotTableSnapshot.Source.Range(sheetName, normalizeArea(area));
    }

    String sourceName = requireNonBlank(worksheetSource.getName(), "Pivot source name is missing.");
    List<Name> namedRanges = matchingNamedRanges(workbook, sourceName, sheetName);
    if (namedRanges.size() == 1) {
      AreaReference area = namedRangeArea(namedRanges.getFirst());
      return new ExcelPivotTableSnapshot.Source.NamedRange(
          sourceName,
          sourceSheetName(area, namedRanges.getFirst(), sheetName),
          normalizeArea(area));
    }
    if (namedRanges.size() > 1) {
      throw new IllegalArgumentException(
          "Pivot source name '"
              + sourceName
              + "' is ambiguous because multiple matching named ranges exist.");
    }

    XSSFTable table = tableByName(workbook, sourceName, sheetName);
    if (table != null) {
      return new ExcelPivotTableSnapshot.Source.Table(
          sourceName,
          table.getSheetName(),
          normalizeArea(
              new AreaReference(
                  table.getStartCellReference(),
                  table.getEndCellReference(),
                  SpreadsheetVersion.EXCEL2007)));
    }
    throw new IllegalArgumentException(
        "Pivot source name '"
            + sourceName
            + "' does not resolve to an existing named range or table.");
  }

  private ResolvedAuthoringSource resolveAuthoringSource(
      ExcelWorkbook workbook, ExcelPivotTableDefinition.Source source) {
    return switch (source) {
      case ExcelPivotTableDefinition.Source.Range range -> {
        XSSFSheet sheet = requiredSheet(workbook, range.sheetName());
        AreaReference area = contiguousArea(range.range(), "range");
        yield ResolvedAuthoringSource.range(sheet, area);
      }
      case ExcelPivotTableDefinition.Source.NamedRange namedRange -> {
        List<Name> matches = matchingNamedRanges(workbook.xssfWorkbook(), namedRange.name(), null);
        if (matches.isEmpty()) {
          throw new IllegalArgumentException(
              "pivot source named range not found: " + namedRange.name());
        }
        if (matches.size() > 1) {
          throw new IllegalArgumentException(
              "pivot source named range is ambiguous across workbook and sheet scope: "
                  + namedRange.name());
        }
        Name resolved = matches.getFirst();
        AreaReference area = namedRangeArea(resolved);
        XSSFSheet sheet = requiredSheet(workbook, sourceSheetName(area, resolved, null));
        yield ResolvedAuthoringSource.namedRange(sheet, area, resolved);
      }
      case ExcelPivotTableDefinition.Source.Table table -> {
        XSSFTable resolved = requiredTableByName(workbook, table.name());
        XSSFSheet sheet = requiredSheet(workbook, resolved.getSheetName());
        AreaReference area =
            new AreaReference(
                resolved.getStartCellReference(),
                resolved.getEndCellReference(),
                SpreadsheetVersion.EXCEL2007);
        yield ResolvedAuthoringSource.table(sheet, area, resolved);
      }
    };
  }

  private XSSFTable requiredTableByName(ExcelWorkbook workbook, String name) {
    XSSFTable table = tableByName(workbook.xssfWorkbook(), name, null);
    if (table == null) {
      throw new IllegalArgumentException("pivot source table not found: " + name);
    }
    return table;
  }

  private XSSFTable tableByName(XSSFWorkbook workbook, String name, String preferredSheetName) {
    String expected = name.toUpperCase(Locale.ROOT);
    XSSFTable preferred = null;
    XSSFTable fallback = null;
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
      for (XSSFTable table : sheet.getTables()) {
        if (!Objects.requireNonNullElse(table.getName(), "")
            .toUpperCase(Locale.ROOT)
            .equals(expected)) {
          continue;
        }
        if (preferredSheetName != null && preferredSheetName.equals(sheet.getSheetName())) {
          if (preferred != null) {
            throw new IllegalArgumentException(
                "Multiple tables named '"
                    + name
                    + "' exist on sheet '"
                    + preferredSheetName
                    + "'.");
          }
          preferred = table;
          continue;
        }
        if (fallback != null) {
          throw new IllegalArgumentException(
              "Multiple tables named '" + name + "' exist in the workbook.");
        }
        fallback = table;
      }
    }
    return preferred != null ? preferred : fallback;
  }

  private List<Name> matchingNamedRanges(XSSFWorkbook workbook, String name, String sheetNameHint) {
    List<Name> matches = new ArrayList<>();
    for (Name candidate : workbook.getAllNames()) {
      if (!Objects.requireNonNullElse(candidate.getNameName(), "").equalsIgnoreCase(name)) {
        continue;
      }
      String candidateSheetName =
          candidate.getSheetIndex() < 0 ? null : workbook.getSheetName(candidate.getSheetIndex());
      if (sheetNameHint == null
          || candidateSheetName == null
          || candidateSheetName.equals(sheetNameHint)) {
        matches.add(candidate);
      }
    }
    return List.copyOf(matches);
  }

  private AreaReference namedRangeArea(Name namedRange) {
    String formula = Objects.requireNonNullElse(namedRange.getRefersToFormula(), "");
    if (formula.isBlank()) {
      throw new IllegalArgumentException(
          "Named range '" + namedRange.getNameName() + "' does not refer to a range.");
    }
    try {
      return new AreaReference(
          ExcelNamedRangeTargets.normalizeAreaFormulaForPoi(formula), SpreadsheetVersion.EXCEL2007);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "Named range '" + namedRange.getNameName() + "' is not a contiguous range source.",
          exception);
    }
  }

  private String sourceSheetName(AreaReference area, Name namedRange, String fallbackSheetName) {
    String areaSheetName = area.getFirstCell().getSheetName();
    if (areaSheetName != null && !areaSheetName.isBlank()) {
      return areaSheetName;
    }
    if (namedRange.getSheetIndex() >= 0) {
      return namedRange.getSheetName();
    }
    if (fallbackSheetName != null && !fallbackSheetName.isBlank()) {
      return fallbackSheetName;
    }
    throw new IllegalArgumentException(
        "Named range '"
            + namedRange.getNameName()
            + "' does not resolve to a concrete source sheet.");
  }

  private SourceColumns sourceColumns(XSSFSheet sheet, AreaReference area, String description) {
    CellReference firstCell = area.getFirstCell();
    CellReference lastCell = area.getLastCell();
    if (firstCell.getRow() >= lastCell.getRow()) {
      throw new IllegalArgumentException(
          "pivot source " + description + " must include a header row plus at least one data row");
    }
    Row headerRow = sheet.getRow(firstCell.getRow());
    if (headerRow == null) {
      throw new IllegalArgumentException(
          "pivot source " + description + " is missing its header row");
    }
    List<SourceColumn> columns = new ArrayList<>();
    Set<String> unique = new LinkedHashSet<>();
    for (int columnIndex = firstCell.getCol(); columnIndex <= lastCell.getCol(); columnIndex++) {
      var cell = headerRow.getCell(columnIndex);
      if (cell == null
          || cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING
          || cell.getStringCellValue().isBlank()) {
        throw new IllegalArgumentException(
            "pivot source "
                + description
                + " header row must contain non-blank text cells for every source column");
      }
      String name = cell.getStringCellValue();
      String key = name.toUpperCase(Locale.ROOT);
      if (!unique.add(key)) {
        throw new IllegalArgumentException(
            "pivot source "
                + description
                + " header row must be unique case-insensitively: "
                + name);
      }
      columns.add(new SourceColumn(name, columnIndex - firstCell.getCol()));
    }
    return new SourceColumns(List.copyOf(columns));
  }

  private List<String> cacheFieldNames(XSSFPivotTable pivotTable) {
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

  private ExcelPivotTableSnapshot.Field sourceField(
      List<String> sourceColumnNames, int sourceColumnIndex) {
    return new ExcelPivotTableSnapshot.Field(
        sourceColumnIndex, sourceColumnName(sourceColumnNames, sourceColumnIndex));
  }

  private String sourceColumnName(List<String> sourceColumnNames, int sourceColumnIndex) {
    if (sourceColumnIndex < 0 || sourceColumnIndex >= sourceColumnNames.size()) {
      throw new IllegalArgumentException(
          "Pivot field references source column index "
              + sourceColumnIndex
              + " but only "
              + sourceColumnNames.size()
              + " cache fields are present.");
    }
    return sourceColumnNames.get(sourceColumnIndex);
  }

  private ExcelPivotDataConsolidateFunction fromSubtotal(int subtotalValue) {
    for (DataConsolidateFunction function : DataConsolidateFunction.values()) {
      if (function.getValue() == subtotalValue) {
        return ExcelPivotDataConsolidateFunction.fromPoiFunction(function);
      }
    }
    throw new IllegalArgumentException(
        "Unsupported pivot data consolidate function value: " + subtotalValue);
  }

  private String numberFormat(XSSFWorkbook workbook, Long numFmtId) {
    if (numFmtId == null) {
      return null;
    }
    short formatIndex = numFmtId > Short.MAX_VALUE ? Short.MAX_VALUE : (short) numFmtId.longValue();
    String format = workbook.createDataFormat().getFormat(formatIndex);
    return format == null || format.isBlank() ? null : format;
  }

  private XSSFPivotCacheDefinition requiredCacheDefinition(XSSFPivotTable pivotTable) {
    XSSFPivotCacheDefinition cacheDefinition = cacheDefinition(pivotTable);
    if (cacheDefinition == null) {
      throw new IllegalArgumentException("Pivot table is missing its cache definition relation.");
    }
    return cacheDefinition;
  }

  private XSSFPivotCacheDefinition cacheDefinition(XSSFPivotTable pivotTable) {
    try {
      return pivotTable.getPivotCacheDefinition();
    } catch (RuntimeException exception) {
      return null;
    }
  }

  private XSSFPivotCacheRecords cacheRecords(XSSFPivotCacheDefinition cacheDefinition) {
    return firstRelation(cacheDefinition, XSSFPivotCacheRecords.class);
  }

  private <T extends POIXMLDocumentPart> T firstRelation(
      POIXMLDocumentPart parent, Class<T> relationType) {
    for (POIXMLDocumentPart relation : parent.getRelations()) {
      if (relationType.isInstance(relation)) {
        return relationType.cast(relation);
      }
    }
    return null;
  }

  private CTPivotCache workbookPivotCache(XSSFWorkbook workbook, long cacheId) {
    for (CTPivotCache cache : workbookPivotCaches(workbook)) {
      if (cache.getCacheId() == cacheId) {
        return cache;
      }
    }
    return null;
  }

  private List<CTPivotCache> workbookPivotCaches(XSSFWorkbook workbook) {
    if (!workbook.getCTWorkbook().isSetPivotCaches()) {
      return List.of();
    }
    return List.of(workbook.getCTWorkbook().getPivotCaches().getPivotCacheArray());
  }

  private PivotLocation safeLocation(PivotHandle handle) {
    String rawRange = rawLocationRange(handle);
    if (rawRange == null) {
      return null;
    }
    try {
      AreaReference area = new AreaReference(rawRange, SpreadsheetVersion.EXCEL2007);
      String locationRange = normalizeArea(area);
      return new PivotLocation(
          new CellReference(area.getFirstCell().getRow(), area.getFirstCell().getCol())
              .formatAsString(),
          locationRange);
    } catch (RuntimeException exception) {
      return null;
    }
  }

  private String rawLocationRange(PivotHandle handle) {
    CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
    if (definition.getLocation() == null) {
      return null;
    }
    return definition.getLocation().getRef();
  }

  private String resolvedName(PivotHandle handle) {
    String actualName = actualName(handle);
    if (actualName != null) {
      return actualName;
    }
    PivotLocation location = safeLocation(handle);
    if (location != null) {
      return syntheticName(handle.sheetName(), location.topLeftAddress());
    }
    return SYNTHETIC_PREFIX
        + sanitize(handle.sheetName())
        + "_"
        + (handle.sheetIndex() + 1)
        + "_"
        + (handle.ordinalOnSheet() + 1);
  }

  private String actualName(PivotHandle handle) {
    String name = handle.table().getCTPivotTableDefinition().getName();
    return name == null || name.isBlank() ? null : name;
  }

  private String syntheticName(String sheetName, String topLeftAddress) {
    return SYNTHETIC_PREFIX + sanitize(sheetName) + "_" + sanitize(topLeftAddress);
  }

  private String sanitize(String value) {
    StringBuilder builder = new StringBuilder(value.length());
    for (int index = 0; index < value.length(); index++) {
      char candidate = value.charAt(index);
      if (Character.isLetterOrDigit(candidate) || candidate == '_') {
        builder.append(candidate);
      } else {
        builder.append('_');
      }
    }
    return builder.toString();
  }

  private AreaReference contiguousArea(String rawRange, String fieldName) {
    try {
      return new AreaReference(rawRange, SpreadsheetVersion.EXCEL2007);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          fieldName + " must be a contiguous A1-style range", exception);
    }
  }

  private String normalizeArea(AreaReference area) {
    CellReference first =
        new CellReference(area.getFirstCell().getRow(), area.getFirstCell().getCol(), false, false);
    CellReference last =
        new CellReference(area.getLastCell().getRow(), area.getLastCell().getCol(), false, false);
    boolean singleCell = first.getRow() == last.getRow() && first.getCol() == last.getCol();
    return singleCell
        ? first.formatAsString()
        : first.formatAsString() + ":" + last.formatAsString();
  }

  private XSSFSheet requiredSheet(ExcelWorkbook workbook, String sheetName) {
    XSSFSheet sheet = workbook.xssfWorkbook().getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  private String requireNonBlank(String value, String message) {
    if (value == null || value.isBlank()) {
      throw new IllegalArgumentException(message);
    }
    return value;
  }

  private String nonBlankOrDefault(String value, String defaultValue) {
    return value == null || value.isBlank() ? defaultValue : value;
  }

  private String normalizedResolvedName(PivotHandle handle) {
    return resolvedName(handle).toUpperCase(Locale.ROOT);
  }

  /** Supported authored pivot-source identities after workbook lookup resolves real POI objects. */
  private enum ResolvedAuthoringSourceKind {
    RANGE,
    NAMED_RANGE,
    TABLE
  }

  private record ResolvedAuthoringSource(
      ResolvedAuthoringSourceKind kind,
      XSSFSheet sheet,
      AreaReference area,
      String description,
      Name namedRange,
      XSSFTable table) {
    private ResolvedAuthoringSource {
      Objects.requireNonNull(kind, "kind must not be null");
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(area, "area must not be null");
      Objects.requireNonNull(description, "description must not be null");
    }

    private static ResolvedAuthoringSource range(XSSFSheet sheet, AreaReference area) {
      return new ResolvedAuthoringSource(
          ResolvedAuthoringSourceKind.RANGE,
          sheet,
          area,
          sheet.getSheetName() + "!" + area.formatAsString(),
          null,
          null);
    }

    private static ResolvedAuthoringSource namedRange(
        XSSFSheet sheet, AreaReference area, Name namedRange) {
      return new ResolvedAuthoringSource(
          ResolvedAuthoringSourceKind.NAMED_RANGE,
          sheet,
          area,
          "named range " + namedRange.getNameName(),
          namedRange,
          null);
    }

    private static ResolvedAuthoringSource table(
        XSSFSheet sheet, AreaReference area, XSSFTable table) {
      return new ResolvedAuthoringSource(
          ResolvedAuthoringSourceKind.TABLE, sheet, area, "table " + table.getName(), null, table);
    }
  }

  private record ColumnAxisSnapshot(
      List<ExcelPivotTableSnapshot.Field> columnLabels, boolean valuesAxisOnColumns) {}

  private record SourceColumn(String name, int relativeIndex) {}

  private record SourceColumns(List<SourceColumn> columns) {
    private SourceColumns {
      columns = List.copyOf(columns);
    }

    int relativeIndex(String name) {
      String expected = name.toUpperCase(Locale.ROOT);
      for (SourceColumn column : columns) {
        if (column.name().toUpperCase(Locale.ROOT).equals(expected)) {
          return column.relativeIndex();
        }
      }
      throw new IllegalArgumentException("pivot source column not found: " + name);
    }
  }

  private record PivotHandle(
      int sheetIndex,
      int ordinalOnSheet,
      String sheetName,
      XSSFSheet sheet,
      XSSFPivotTable table) {}

  private record PivotLocation(String topLeftAddress, String locationRange) {}
}
