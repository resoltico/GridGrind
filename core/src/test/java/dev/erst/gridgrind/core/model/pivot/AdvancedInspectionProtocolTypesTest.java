package dev.erst.gridgrind.core.model.pivot;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.WorkbookAnalysisResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
            new StepId("security"),
            new OoxmlPackageSecurityReport(new OoxmlEncryptionReport.None(), List.of()));

    assertEquals("GET_PACKAGE_SECURITY", query.queryType());
    assertEquals("security", result.stepId().value());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookInspectionResult.PackageSecurityResult(new StepId(" "), result.security()));
  }

  @Test
  void pivotInspectionTypesRetainAnalysisPayloads() {
    WorkbookAssetInspectionResult.PivotTablesResult pivots =
        new WorkbookAssetInspectionResult.PivotTablesResult(
            new StepId("pivots"),
            List.of(
                new PivotTableReport.Supported(
                    "Sales Pivot 2026",
                    new SheetName("Report"),
                    new PivotTableReport.Anchor(new CellAddress("C5"), new RangeLabel("C5:G9")),
                    new PivotTableReport.Source.Range(
                        new SheetName("Data"), new RangeLabel("A1:D5")),
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
                    new AnalysisLocationReport.Sheet(new SheetLabel("Report")),
                    List.of("_GG_PIVOT_Report_A3"))));
    WorkbookAnalysisResult.PivotTableHealthResult pivotHealth =
        new WorkbookAnalysisResult.PivotTableHealthResult(new StepId("pivot-health"), health);

    assertEquals("Sales Pivot 2026", pivots.pivotTables().getFirst().name());
    assertEquals(1, pivotHealth.analysis().checkedPivotTableCount());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAssetInspectionResult.PivotTablesResult(new StepId("pivots"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysisResult.PivotTableHealthResult(new StepId("pivot-health"), null));
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
