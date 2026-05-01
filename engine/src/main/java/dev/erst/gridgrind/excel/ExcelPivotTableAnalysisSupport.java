package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

/** Owns pivot-table integrity and duplicate-name analysis. */
final class ExcelPivotTableAnalysisSupport {
  private ExcelPivotTableAnalysisSupport() {}

  static List<WorkbookAnalysis.AnalysisFinding> duplicateNameFindings(List<PivotHandle> handles) {
    Set<String> seenNames = new LinkedHashSet<>();
    Set<String> duplicateNames = new LinkedHashSet<>();
    for (PivotHandle handle : handles) {
      String normalizedName = ExcelPivotTableIdentitySupport.normalizedResolvedName(handle);
      if (!seenNames.add(normalizedName)) {
        duplicateNames.add(normalizedName);
      }
    }
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (PivotHandle handle : handles) {
      if (!duplicateNames.contains(ExcelPivotTableIdentitySupport.normalizedResolvedName(handle))) {
        continue;
      }
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode.PIVOT_TABLE_DUPLICATE_NAME,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
              handle,
              "Pivot table name is not unique",
              "Multiple pivot tables share the same case-insensitive name, so exact-name selection is ambiguous.",
              List.of(ExcelPivotTableIdentitySupport.resolvedName(handle))));
    }
    return List.copyOf(findings);
  }

  static List<WorkbookAnalysis.AnalysisFinding> pivotTableHealthFindings(
      XSSFWorkbook workbook, PivotHandle handle) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    if (ExcelPivotTableIdentitySupport.actualName(handle) == null) {
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.WARNING,
              handle,
              "Pivot table name is missing",
              "The pivot table does not persist a name, so GridGrind assigned a synthetic identifier for readback.",
              List.of(ExcelPivotTableIdentitySupport.resolvedName(handle))));
    }

    PivotLocation location = ExcelPivotTableIdentitySupport.safeLocation(handle).orElse(null);
    if (location == null) {
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                  .PIVOT_TABLE_UNSUPPORTED_DETAIL,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
              handle,
              "Pivot table location is malformed",
              "The pivot table location range could not be parsed.",
              List.of(ExcelPivotTableIdentitySupport.rawLocationRange(handle))));
      return List.copyOf(findings);
    }

    XSSFPivotCacheDefinition cacheDefinition =
        ExcelPivotTableSnapshotSupport.cacheDefinition(handle.table());
    if (cacheDefinition == null) {
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                  .PIVOT_TABLE_MISSING_CACHE_DEFINITION,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
              handle,
              "Pivot table is missing its cache definition relation",
              "The pivot table part no longer points at a pivot cache definition.",
              List.of(location.locationRange())));
      return List.copyOf(findings);
    }

    if (ExcelPivotTableSnapshotSupport.cacheRecords(cacheDefinition) == null) {
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                  .PIVOT_TABLE_MISSING_CACHE_RECORDS,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
              handle,
              "Pivot table is missing its cache records relation",
              "The pivot cache definition does not point at pivot cache records.",
              List.of(location.locationRange())));
    }

    CTPivotTableDefinition definition = handle.table().getCTPivotTableDefinition();
    if (ExcelPivotTableSnapshotSupport.workbookPivotCache(workbook, definition.getCacheId())
        == null) {
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                  .PIVOT_TABLE_MISSING_WORKBOOK_CACHE,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
              handle,
              "Pivot table cache is not registered in workbook metadata",
              "The pivot table cacheId is missing from workbook.xml pivotCaches.",
              List.of(Long.toString(definition.getCacheId()))));
    }

    ExcelPivotTableSnapshot snapshot = ExcelPivotTableSnapshotSupport.snapshot(workbook, handle);
    if (snapshot instanceof ExcelPivotTableSnapshot.Unsupported unsupported) {
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                  .PIVOT_TABLE_UNSUPPORTED_DETAIL,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.WARNING,
              handle,
              "Pivot table contains unsupported detail",
              unsupported.detail(),
              List.of(unsupported.anchor().locationRange())));
    }

    try {
      ExcelPivotTableSnapshotSupport.snapshotSource(workbook, handle.table());
    } catch (RuntimeException exception) {
      findings.add(
          finding(
              dev.erst.gridgrind.excel.foundation.AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE,
              dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
              handle,
              "Pivot table source is broken",
              Objects.requireNonNullElse(
                  exception.getMessage(), "The pivot source no longer resolves cleanly."),
              List.of(location.locationRange())));
    }
    return List.copyOf(new ArrayList<>(new LinkedHashSet<>(findings)));
  }

  static WorkbookAnalysis.AnalysisFinding finding(
      dev.erst.gridgrind.excel.foundation.AnalysisFindingCode code,
      dev.erst.gridgrind.excel.foundation.AnalysisSeverity severity,
      PivotHandle handle,
      String title,
      String message,
      List<String> evidence) {
    PivotLocation location = ExcelPivotTableIdentitySupport.safeLocation(handle).orElse(null);
    WorkbookAnalysis.AnalysisLocation analysisLocation =
        location == null
            ? new WorkbookAnalysis.AnalysisLocation.Sheet(handle.sheetName())
            : new WorkbookAnalysis.AnalysisLocation.Range(
                handle.sheetName(), location.locationRange());
    return new WorkbookAnalysis.AnalysisFinding(
        code, severity, title, message, analysisLocation, List.copyOf(evidence));
  }
}
