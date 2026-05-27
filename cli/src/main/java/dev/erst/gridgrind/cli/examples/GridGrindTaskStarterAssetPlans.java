package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.InspectionSurfaceQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;

/** Published starter plans that require shipped workbook or XML assets. */
final class GridGrindTaskStarterAssetPlans {
  private GridGrindTaskStarterAssetPlans() {}

  static List<TaskStarterPlan> starters() {
    return List.of(auditExistingWorkbookStarter(), customXmlStarter());
  }

  private static TaskStarterPlan auditExistingWorkbookStarter() {
    String taskId = "AUDIT_EXISTING_WORKBOOK";
    String sourceAsset = TaskStarterPlanSupport.taskStarterAsset("workbook-ops-source.xlsx");
    return TaskStarterPlanSupport.assetBackedStarter(
        taskId,
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.ExistingFile(sourceAsset),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExamplePlanSupport.read(
                "read-workbook-summary",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "read-template-comments",
                ExamplePlanSupport.cells("Template", "A1"),
                new SheetIntrospectionQuery.GetComments()),
            ExamplePlanSupport.read(
                "read-drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Template"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            ExamplePlanSupport.read(
                "read-formula-surface",
                ExamplePlanSupport.sheets("Summary"),
                new InspectionSurfaceQuery.GetFormulaSurface()),
            ExamplePlanSupport.read(
                "analyze-workbook-findings",
                ExamplePlanSupport.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())),
        sourceAsset);
  }

  private static TaskStarterPlan customXmlStarter() {
    String taskId = "CUSTOM_XML_WORKFLOW";
    String mappingWorkbook = "custom-xml-assets/custom-xml-mapping.xlsx";
    String updateXml = "custom-xml-assets/custom-xml-update.xml";
    return TaskStarterPlanSupport.assetBackedStarter(
        taskId,
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.ExistingFile(mappingWorkbook),
            ExamplePlanSupport.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExamplePlanSupport.read(
                "read-custom-xml-mappings",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetCustomXmlMappings()),
            ExamplePlanSupport.read(
                "export-custom-xml-before-import",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8")),
            ExamplePlanSupport.step(
                "import-custom-xml",
                ExamplePlanSupport.workbook(),
                new StructuredMutationAction.ImportCustomXmlMapping(
                    new CustomXmlImportInput(
                        new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                        TextSourceInput.utf8File(updateXml)))),
            ExamplePlanSupport.read(
                "read-imported-cells",
                ExamplePlanSupport.cells("Foglio1", "A1", "B1", "C1", "D1"),
                new SheetIntrospectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "export-custom-xml-after-import",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8"))),
        mappingWorkbook,
        updateXml);
  }
}
