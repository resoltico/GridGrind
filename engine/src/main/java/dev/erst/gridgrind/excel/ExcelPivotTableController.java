package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCache;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheRecords;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

/** Reads, writes, and analyzes workbook pivot tables within the POI-supported XSSF surface. */
final class ExcelPivotTableController {
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

    ResolvedAuthoringSource source =
        ExcelPivotTableSourceSupport.resolveAuthoringSource(workbook, definition.source());
    SourceColumns columns =
        ExcelPivotTableSourceSupport.sourceColumns(
            source.sheet(), source.area(), source.description());
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
          ExcelPivotDataPoiBridge.toPoi(dataField.function()),
          columns.relativeIndex(dataField.sourceColumnName()),
          dataField.displayName(),
          dataField.valueFormat());
    }
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet name. */
  void deletePivotTable(ExcelWorkbook workbook, String name, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String validatedName = ExcelPivotTableNaming.validateName(name);
    dev.erst.gridgrind.excel.foundation.ExcelSheetNames.requireValid(sheetName, "sheetName");

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

  XSSFPivotTable createPivotTable(
      ExcelWorkbook workbook,
      ExcelPivotTableDefinition definition,
      ResolvedAuthoringSource source,
      CellReference anchor) {
    var sheet = ExcelPivotTableSourceSupport.requiredSheet(workbook, definition.sheetName());
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

  void normalizeCacheId(XSSFWorkbook workbook, XSSFPivotTable pivotTable) {
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

  List<PivotHandle> selectHandles(ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    List<PivotHandle> all = allPivotTables(workbook);
    return switch (selection) {
      case ExcelPivotTableSelection.All ignored -> all;
      case ExcelPivotTableSelection.ByNames byNames -> selectHandlesByName(all, byNames.names());
    };
  }

  List<PivotHandle> selectHandlesByName(List<PivotHandle> handles, List<String> names) {
    List<PivotHandle> selected = new ArrayList<>();
    for (String name : names) {
      PivotHandle handle = findPivotHandleByName(handles, name);
      if (handle != null) {
        selected.add(handle);
      }
    }
    return List.copyOf(selected);
  }

  PivotHandle pivotByName(ExcelWorkbook workbook, String name) {
    return findPivotHandleByName(allPivotTables(workbook), name);
  }

  PivotHandle findPivotHandleByName(List<PivotHandle> handles, String name) {
    String expected = ExcelPivotTableNaming.validateName(name).toUpperCase(Locale.ROOT);
    for (PivotHandle handle : handles) {
      if (ExcelPivotTableIdentitySupport.resolvedName(handle)
          .toUpperCase(Locale.ROOT)
          .equals(expected)) {
        return handle;
      }
    }
    return null;
  }

  List<PivotHandle> allPivotTables(ExcelWorkbook workbook) {
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

  List<WorkbookAnalysis.AnalysisFinding> duplicateNameFindings(List<PivotHandle> handles) {
    return ExcelPivotTableAnalysisSupport.duplicateNameFindings(handles);
  }

  List<WorkbookAnalysis.AnalysisFinding> pivotTableHealthFindings(
      XSSFWorkbook workbook, PivotHandle handle) {
    return ExcelPivotTableAnalysisSupport.pivotTableHealthFindings(workbook, handle);
  }

  WorkbookAnalysis.AnalysisFinding finding(
      dev.erst.gridgrind.excel.foundation.AnalysisFindingCode code,
      dev.erst.gridgrind.excel.foundation.AnalysisSeverity severity,
      PivotHandle handle,
      String title,
      String message,
      List<String> evidence) {
    return ExcelPivotTableAnalysisSupport.finding(code, severity, handle, title, message, evidence);
  }

  void deletePivotHandle(ExcelWorkbook workbook, PivotHandle handle) {
    ExcelPivotTableLifecycleSupport.deletePivotHandle(
        workbook, handle, allPivotTables(workbook), poiRelationRemover);
  }

  void removeWorkbookPivotCacheRegistration(
      XSSFWorkbook workbook, long cacheId, String relationId) {
    ExcelPivotTableLifecycleSupport.removeWorkbookPivotCacheRegistration(
        workbook, cacheId, relationId);
  }

  boolean cacheDefinitionShared(
      ExcelWorkbook workbook, PivotHandle current, XSSFPivotCacheDefinition cacheDefinition) {
    return ExcelPivotTableLifecycleSupport.cacheDefinitionShared(
        workbook, current, cacheDefinition, allPivotTables(workbook));
  }

  boolean removePoiRelation(POIXMLDocumentPart parent, POIXMLDocumentPart child) {
    return ExcelPivotTableLifecycleSupport.removePoiRelation(parent, child, poiRelationRemover);
  }

  void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg,
      org.apache.poi.openxml4j.opc.PackagePartName partName) {
    ExcelPivotTableLifecycleSupport.cleanupPackagePartIfUnused(pkg, partName);
  }

  void primePivotTableAllocator(XSSFWorkbook workbook, XSSFPivotTable allocationSentinel) {
    ExcelPivotTableLifecycleSupport.primePivotTableAllocator(workbook, allocationSentinel);
  }

  void rebuildPivotTableRegistry(XSSFWorkbook workbook) {
    ExcelPivotTableLifecycleSupport.rebuildPivotTableRegistry(workbook);
  }

  int pivotTableIdHighWaterMark(XSSFWorkbook workbook) {
    return ExcelPivotTableLifecycleSupport.pivotTableIdHighWaterMark(workbook);
  }

  int packagePartIndex(POIXMLDocumentPart part, String prefix) {
    return ExcelPivotTableLifecycleSupport.packagePartIndex(part, prefix);
  }

  int packagePartIndex(PackagePart part, String prefix) {
    return ExcelPivotTableLifecycleSupport.packagePartIndex(part, prefix);
  }

  ExcelPivotTableSnapshot snapshot(XSSFWorkbook workbook, PivotHandle handle) {
    return ExcelPivotTableSnapshotSupport.snapshot(workbook, handle);
  }

  ColumnAxisSnapshot snapshotColumnLabels(
      CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    return ExcelPivotTableSnapshotSupport.snapshotColumnLabels(definition, sourceColumnNames);
  }

  List<ExcelPivotTableSnapshot.Field> snapshotFields(
      CTField[] fields, List<String> sourceColumnNames) {
    return ExcelPivotTableSnapshotSupport.snapshotFields(fields, sourceColumnNames);
  }

  List<ExcelPivotTableSnapshot.Field> snapshotPageFields(
      CTPageField[] pageFields, List<String> sourceColumnNames) {
    return ExcelPivotTableSnapshotSupport.snapshotPageFields(pageFields, sourceColumnNames);
  }

  List<ExcelPivotTableSnapshot.DataField> snapshotDataFields(
      XSSFWorkbook workbook, CTPivotTableDefinition definition, List<String> sourceColumnNames) {
    return ExcelPivotTableSnapshotSupport.snapshotDataFields(
        workbook, definition, sourceColumnNames);
  }

  ExcelPivotTableSnapshot.Unsupported unsupportedSnapshot(
      PivotHandle handle, String name, ExcelPivotTableSnapshot.Anchor anchor, String detail) {
    return ExcelPivotTableSnapshotSupport.unsupportedSnapshot(handle, name, anchor, detail);
  }

  ExcelPivotTableSnapshot.Source snapshotSource(XSSFWorkbook workbook, XSSFPivotTable pivotTable) {
    return ExcelPivotTableSnapshotSupport.snapshotSource(workbook, pivotTable);
  }

  List<String> cacheFieldNames(XSSFPivotTable pivotTable) {
    return ExcelPivotTableSnapshotSupport.cacheFieldNames(pivotTable);
  }

  ExcelPivotTableSnapshot.Field sourceField(List<String> sourceColumnNames, int sourceColumnIndex) {
    return ExcelPivotTableSnapshotSupport.sourceField(sourceColumnNames, sourceColumnIndex);
  }

  ExcelPivotDataConsolidateFunction fromSubtotal(int subtotalValue) {
    return ExcelPivotTableSnapshotSupport.fromSubtotal(subtotalValue);
  }

  String numberFormat(XSSFWorkbook workbook, Long numFmtId) {
    return ExcelPivotTableSnapshotSupport.numberFormat(workbook, numFmtId);
  }

  XSSFPivotCacheDefinition requiredCacheDefinition(XSSFPivotTable pivotTable) {
    return ExcelPivotTableSnapshotSupport.requiredCacheDefinition(pivotTable);
  }

  XSSFPivotCacheDefinition cacheDefinition(XSSFPivotTable pivotTable) {
    return ExcelPivotTableSnapshotSupport.cacheDefinition(pivotTable);
  }

  XSSFPivotCacheRecords cacheRecords(XSSFPivotCacheDefinition cacheDefinition) {
    return ExcelPivotTableSnapshotSupport.cacheRecords(cacheDefinition);
  }

  <T extends POIXMLDocumentPart> T firstRelation(POIXMLDocumentPart parent, Class<T> relationType) {
    return ExcelPivotTableSnapshotSupport.firstRelation(parent, relationType);
  }

  CTPivotCache workbookPivotCache(XSSFWorkbook workbook, long cacheId) {
    return ExcelPivotTableSnapshotSupport.workbookPivotCache(workbook, cacheId);
  }

  List<CTPivotCache> workbookPivotCaches(XSSFWorkbook workbook) {
    return ExcelPivotTableSnapshotSupport.workbookPivotCaches(workbook);
  }
}
