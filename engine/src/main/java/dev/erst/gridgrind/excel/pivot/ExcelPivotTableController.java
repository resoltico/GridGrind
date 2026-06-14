package dev.erst.gridgrind.excel.pivot;

import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.PoiRelationRemoval;
import dev.erst.gridgrind.excel.WorkbookAnalysis;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCache;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;

/** Reads, writes, and analyzes workbook pivot tables within the POI-supported XSSF surface. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelPivotTableController {
  private final BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover;

  public ExcelPivotTableController() {
    this(PoiRelationRemoval.defaultRemover());
  }

  public ExcelPivotTableController(
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    this.poiRelationRemover =
        Objects.requireNonNull(poiRelationRemover, "poiRelationRemover must not be null");
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  public void setPivotTable(ExcelWorkbook workbook, ExcelPivotTableDefinition definition) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    Optional<PivotHandle> existing = pivotByName(workbook, definition.name());
    if (existing.isPresent()
        && !existing.orElseThrow().sheetName().equals(definition.sheetName())) {
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

    if (existing.isPresent()) {
      ExcelPivotTableLifecycleSupport.deletePivotHandle(
          workbook, existing.orElseThrow(), allPivotTables(workbook), poiRelationRemover);
    }

    ExcelPivotTableLifecycleSupport.primePivotTableAllocator(
        workbook.xssfWorkbook(), existing.map(PivotHandle::table));
    try {
      XSSFPivotTable pivotTable = createPivotTable(workbook, definition, source, anchor);
      normalizeCacheId(workbook.xssfWorkbook(), pivotTable);
      pivotTable.getCTPivotTableDefinition().setName(definition.name());
      applyPivotFields(pivotTable, definition, columns);
    } finally {
      ExcelPivotTableLifecycleSupport.rebuildPivotTableRegistry(workbook.xssfWorkbook());
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
          dataField.valueFormat().orElse(null));
    }
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet name. */
  public void deletePivotTable(ExcelWorkbook workbook, String name, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String validatedName = ExcelPivotTableNaming.validateName(name);
    dev.erst.gridgrind.excel.foundation.ExcelSheetNames.requireValid(sheetName, "sheetName");

    Optional<PivotHandle> handle = pivotByName(workbook, validatedName);
    if (handle.isEmpty() || !handle.orElseThrow().sheetName().equals(sheetName)) {
      throw new IllegalArgumentException(
          "pivot table not found on expected sheet: " + validatedName + "@" + sheetName);
    }
    ExcelPivotTableLifecycleSupport.deletePivotHandle(
        workbook, handle.orElseThrow(), allPivotTables(workbook), poiRelationRemover);
  }

  /** Returns factual pivot-table metadata selected by workbook-global name or all pivots. */
  public List<ExcelPivotTableSnapshot> pivotTables(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<PivotHandle> handles = selectHandles(workbook, selection);
    List<ExcelPivotTableSnapshot> snapshots = new ArrayList<>(handles.size());
    for (PivotHandle handle : handles) {
      snapshots.add(ExcelPivotTableSnapshotSupport.snapshot(workbook.xssfWorkbook(), handle));
    }
    return List.copyOf(snapshots);
  }

  /** Returns integrity findings for the selected pivot-table set. */
  public List<WorkbookAnalysis.AnalysisFinding> pivotTableHealthFindings(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<PivotHandle> handles = selectHandles(workbook, selection);
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (PivotHandle handle : handles) {
      findings.addAll(
          ExcelPivotTableAnalysisSupport.pivotTableHealthFindings(workbook.xssfWorkbook(), handle));
    }
    findings.addAll(ExcelPivotTableAnalysisSupport.duplicateNameFindings(handles));
    return List.copyOf(new ArrayList<>(new LinkedHashSet<>(findings)));
  }

  public XSSFPivotTable createPivotTable(
      ExcelWorkbook workbook,
      ExcelPivotTableDefinition definition,
      ResolvedAuthoringSource source,
      CellReference anchor) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");

    var sheet = ExcelPivotTableSourceSupport.requiredSheet(workbook, definition.sheetName());
    ExcelPivotTableSourceSupport.sourceColumns(source.sheet(), source.area(), source.description());
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
    Optional<XSSFPivotCacheDefinition> cacheDefinition =
        ExcelPivotTableSnapshotSupport.cacheDefinition(pivotTable);
    String currentRelationId = cacheDefinition.map(workbook::getRelationId).orElse(null);
    long currentId = pivotCache.getCTPivotCache().getCacheId();
    long maxOtherId = 0L;
    boolean duplicate = false;
    for (CTPivotCache cache : ExcelPivotTableSnapshotSupport.workbookPivotCaches(workbook)) {
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

  public List<PivotHandle> selectHandles(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    List<PivotHandle> all = allPivotTables(workbook);
    return switch (selection) {
      case ExcelPivotTableSelection.All ignored -> all;
      case ExcelPivotTableSelection.ByNames byNames -> selectHandlesByName(all, byNames.names());
    };
  }

  public List<PivotHandle> selectHandlesByName(List<PivotHandle> handles, List<String> names) {
    List<PivotHandle> selected = new ArrayList<>();
    for (String name : names) {
      findPivotHandleByName(handles, name).ifPresent(selected::add);
    }
    return List.copyOf(selected);
  }

  public Optional<PivotHandle> pivotByName(ExcelWorkbook workbook, String name) {
    return findPivotHandleByName(allPivotTables(workbook), name);
  }

  public Optional<PivotHandle> findPivotHandleByName(List<PivotHandle> handles, String name) {
    String expected = ExcelPivotTableNaming.validateName(name).toUpperCase(Locale.ROOT);
    for (PivotHandle handle : handles) {
      if (ExcelPivotTableIdentitySupport.resolvedName(handle)
          .toUpperCase(Locale.ROOT)
          .equals(expected)) {
        return Optional.of(handle);
      }
    }
    return Optional.empty();
  }

  public List<PivotHandle> allPivotTables(ExcelWorkbook workbook) {
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

  public Optional<XSSFPivotCacheDefinition> cacheDefinition(XSSFPivotTable pivotTable) {
    return ExcelPivotTableSnapshotSupport.cacheDefinition(pivotTable);
  }

  public void deletePivotHandle(ExcelWorkbook workbook, PivotHandle handle) {
    ExcelPivotTableLifecycleSupport.deletePivotHandle(
        workbook, handle, allPivotTables(workbook), poiRelationRemover);
  }

  public void removeWorkbookPivotCacheRegistration(
      XSSFWorkbook workbook, long cacheId, String relationId) {
    ExcelPivotTableLifecycleSupport.removeWorkbookPivotCacheRegistration(
        workbook, cacheId, relationId);
  }

  public boolean cacheDefinitionShared(
      ExcelWorkbook workbook, PivotHandle current, XSSFPivotCacheDefinition cacheDefinition) {
    return ExcelPivotTableLifecycleSupport.cacheDefinitionShared(
        workbook, current, cacheDefinition, allPivotTables(workbook));
  }

  public boolean removePoiRelation(POIXMLDocumentPart parent, POIXMLDocumentPart child) {
    return ExcelPivotTableLifecycleSupport.removePoiRelation(parent, child, poiRelationRemover);
  }

  public void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg,
      org.apache.poi.openxml4j.opc.PackagePartName partName) {
    ExcelPivotTableLifecycleSupport.cleanupPackagePartIfUnused(pkg, partName);
  }

  public void primePivotTableAllocator(
      XSSFWorkbook workbook, Optional<XSSFPivotTable> allocationSentinel) {
    ExcelPivotTableLifecycleSupport.primePivotTableAllocator(workbook, allocationSentinel);
  }

  public int pivotTableIdHighWaterMark(XSSFWorkbook workbook) {
    return ExcelPivotTableLifecycleSupport.pivotTableIdHighWaterMark(workbook);
  }

  public int packagePartIndex(POIXMLDocumentPart part, String prefix) {
    return ExcelPivotTableLifecycleSupport.packagePartIndex(part, prefix);
  }

  public int packagePartIndex(PackagePart part, String prefix) {
    return ExcelPivotTableLifecycleSupport.packagePartIndex(part, prefix);
  }

  public List<String> cacheFieldNames(XSSFPivotTable pivotTable) {
    return ExcelPivotTableSnapshotSupport.cacheFieldNames(pivotTable);
  }
}
