package dev.erst.gridgrind.excel.pivot;

import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.PoiPackageInspection;
import java.util.List;
import java.util.Optional;
import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheRecords;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;

/** Owns destructive pivot-table lifecycle and package-part cleanup helpers. */
public final class ExcelPivotTableLifecycleSupport {
  private ExcelPivotTableLifecycleSupport() {}

  static void deletePivotHandle(
      ExcelWorkbook workbook,
      PivotHandle handle,
      List<PivotHandle> allPivotTables,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    Optional<XSSFPivotCacheDefinition> cacheDefinition =
        ExcelPivotTableSnapshotSupport.cacheDefinition(handle.table());
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
          "Failed to remove pivot table relation for '"
              + ExcelPivotTableIdentitySupport.resolvedName(handle)
              + "'");
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
            "Failed to remove pivot cache definition relation for '"
                + ExcelPivotTableIdentitySupport.resolvedName(handle)
                + "'");
      }
      cleanupPackagePartIfUnused(workbook.xssfWorkbook().getPackage(), cacheDefinitionPartName);
    }
    rebuildPivotTableRegistry(workbook.xssfWorkbook());
  }

  static void removeWorkbookPivotCacheRegistration(
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

  static boolean cacheDefinitionShared(
      ExcelWorkbook workbook,
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
      Optional<XSSFPivotCacheDefinition> otherCacheDefinition =
          ExcelPivotTableSnapshotSupport.cacheDefinition(handle.table());
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

  static boolean removePoiRelation(
      POIXMLDocumentPart parent,
      POIXMLDocumentPart child,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    return poiRelationRemover.test(parent, child);
  }

  static void cleanupPackagePartIfUnused(
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

  static void primePivotTableAllocator(
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

  static void rebuildPivotTableRegistry(XSSFWorkbook workbook) {
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

  static int pivotTableIdHighWaterMark(XSSFWorkbook workbook) {
    int maximum = 0;
    for (CTPivotCache cache : ExcelPivotTableSnapshotSupport.workbookPivotCaches(workbook)) {
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
        Optional<XSSFPivotCacheDefinition> cacheDefinition =
            ExcelPivotTableSnapshotSupport.cacheDefinition(pivotTable);
        if (cacheDefinition.isPresent()) {
          maximum =
              Math.max(
                  maximum,
                  packagePartIndex(
                      cacheDefinition.orElseThrow(), "/xl/pivotCache/pivotCacheDefinition"));
          Optional<XSSFPivotCacheRecords> cacheRecords =
              ExcelPivotTableSnapshotSupport.cacheRecords(cacheDefinition.orElseThrow());
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

  static int packagePartIndex(POIXMLDocumentPart part, String prefix) {
    if (part.getPackagePart() == null) {
      return 0;
    }
    return packagePartIndex(part.getPackagePart(), prefix);
  }

  static int packagePartIndex(PackagePart part, String prefix) {
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
}
