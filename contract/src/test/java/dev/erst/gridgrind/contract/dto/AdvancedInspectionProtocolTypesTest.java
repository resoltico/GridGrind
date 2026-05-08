package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.WorkbookAnalysisResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct tests for advanced inspection-query and inspection-result DTO families. */
class AdvancedInspectionProtocolTypesTest {
  @Test
  void packageSecurityQueryAndResultUseStepIdentity() {
    WorkbookIntrospectionQuery.GetPackageSecurity query =
        new WorkbookIntrospectionQuery.GetPackageSecurity();
    WorkbookInspectionResult.PackageSecurityResult result =
        new WorkbookInspectionResult.PackageSecurityResult(
            "security",
            new OoxmlPackageSecurityReport(
                new OoxmlEncryptionReport(
                    false,
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty()),
                List.of()));

    assertEquals("GET_PACKAGE_SECURITY", query.queryType());
    assertEquals("security", result.stepId());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookInspectionResult.PackageSecurityResult(" ", result.security()));
  }

  @Test
  void pivotInspectionTypesRetainAnalysisPayloads() {
    WorkbookAssetInspectionResult.PivotTablesResult pivots =
        new WorkbookAssetInspectionResult.PivotTablesResult(
            "pivots",
            List.of(
                new PivotTableReport.Supported(
                    "Sales Pivot 2026",
                    "Report",
                    new PivotTableReport.Anchor("C5", "C5:G9"),
                    new PivotTableReport.Source.Range("Data", "A1:D5"),
                    List.of(new PivotTableReport.Field(0, "Region")),
                    List.of(new PivotTableReport.Field(1, "Stage")),
                    List.of(new PivotTableReport.Field(2, "Owner")),
                    List.of(
                        new PivotTableReport.DataField(
                            3,
                            "Amount",
                            ExcelPivotDataConsolidateFunction.SUM,
                            "Total Amount",
                            Optional.of("#,##0.00"))),
                    true)));
    PivotTableHealthReport health =
        new PivotTableHealthReport(
            1,
            new AnalysisSummaryReport(1, 0, 1, 0),
            List.of(
                new AnalysisFindingReport(
                    AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
                    AnalysisSeverity.WARNING,
                    "Pivot table name is missing",
                    "GridGrind assigned a synthetic identifier for readback.",
                    new AnalysisLocationReport.Sheet("Report"),
                    List.of("_GG_PIVOT_Report_A3"))));
    WorkbookAnalysisResult.PivotTableHealthResult pivotHealth =
        new WorkbookAnalysisResult.PivotTableHealthResult("pivot-health", health);

    assertEquals("Sales Pivot 2026", pivots.pivotTables().getFirst().name());
    assertEquals(1, pivotHealth.analysis().checkedPivotTableCount());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAssetInspectionResult.PivotTablesResult("pivots", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysisResult.PivotTableHealthResult("pivot-health", null));
  }

  @Test
  void workbookProtectionAndNamedRangeQueriesKeepSelectorFamiliesSeparate() {
    WorkbookSelector.Current workbookTarget = new WorkbookSelector.Current();
    NamedRangeSelector.WorkbookScope namedRangeTarget =
        new NamedRangeSelector.WorkbookScope("BudgetTotal");

    assertEquals(
        "GET_WORKBOOK_PROTECTION",
        new WorkbookIntrospectionQuery.GetWorkbookProtection().queryType());
    assertEquals(
        "ANALYZE_NAMED_RANGE_HEALTH",
        new InspectionAnalysisQuery.AnalyzeNamedRangeHealth().queryType());
    assertEquals("BudgetTotal", namedRangeTarget.name());
    assertEquals(WorkbookSelector.Current.class, workbookTarget.getClass());
    assertEquals(PivotTableSelector.All.class, new PivotTableSelector.All().getClass());
  }
}
