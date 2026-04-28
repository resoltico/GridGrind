package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct tests for advanced inspection-query and inspection-result DTO families. */
class AdvancedInspectionProtocolTypesTest {
  @Test
  void packageSecurityQueryAndResultUseStepIdentity() {
    InspectionQuery.GetPackageSecurity query = new InspectionQuery.GetPackageSecurity();
    InspectionResult.PackageSecurityResult result =
        new InspectionResult.PackageSecurityResult(
            "security",
            new OoxmlPackageSecurityReport(
                new OoxmlEncryptionReport(false, null, null, null, null, null, null, null),
                List.of()));

    assertEquals("GET_PACKAGE_SECURITY", query.queryType());
    assertEquals("security", result.stepId());
    assertThrows(
        IllegalArgumentException.class,
        () -> new InspectionResult.PackageSecurityResult(" ", result.security()));
  }

  @Test
  void pivotInspectionTypesRetainAnalysisPayloads() {
    InspectionResult.PivotTablesResult pivots =
        new InspectionResult.PivotTablesResult(
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
                            "#,##0.00")),
                    true)));
    PivotTableHealthReport health =
        new PivotTableHealthReport(
            1,
            new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0),
            List.of(
                new GridGrindResponse.AnalysisFindingReport(
                    AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
                    AnalysisSeverity.WARNING,
                    "Pivot table name is missing",
                    "GridGrind assigned a synthetic identifier for readback.",
                    new GridGrindResponse.AnalysisLocationReport.Sheet("Report"),
                    List.of("_GG_PIVOT_Report_A3"))));
    InspectionResult.PivotTableHealthResult pivotHealth =
        new InspectionResult.PivotTableHealthResult("pivot-health", health);

    assertEquals("Sales Pivot 2026", pivots.pivotTables().getFirst().name());
    assertEquals(1, pivotHealth.analysis().checkedPivotTableCount());
    assertThrows(
        NullPointerException.class, () -> new InspectionResult.PivotTablesResult("pivots", null));
    assertThrows(
        NullPointerException.class,
        () -> new InspectionResult.PivotTableHealthResult("pivot-health", null));
  }

  @Test
  void workbookProtectionAndNamedRangeQueriesKeepSelectorFamiliesSeparate() {
    WorkbookSelector.Current workbookTarget = new WorkbookSelector.Current();
    NamedRangeSelector.WorkbookScope namedRangeTarget =
        new NamedRangeSelector.WorkbookScope("BudgetTotal");

    assertEquals(
        "GET_WORKBOOK_PROTECTION", new InspectionQuery.GetWorkbookProtection().queryType());
    assertEquals(
        "ANALYZE_NAMED_RANGE_HEALTH", new InspectionQuery.AnalyzeNamedRangeHealth().queryType());
    assertEquals("BudgetTotal", namedRangeTarget.name());
    assertEquals(WorkbookSelector.Current.class, workbookTarget.getClass());
    assertEquals(PivotTableSelector.All.class, new PivotTableSelector.All().getClass());
  }
}
