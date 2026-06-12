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
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.ExistingFile(sourceAsset),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExampleSteps.read(
                "read-workbook-summary",
                ExampleSelectors.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            ExampleSteps.read(
                "read-template-comments",
                ExampleSelectors.cells("Template", "A1"),
                new SheetIntrospectionQuery.GetComments()),
            ExampleSteps.read(
                "read-drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Template"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            ExampleSteps.read(
                "read-formula-surface",
                ExampleSelectors.sheets("Summary"),
                new InspectionSurfaceQuery.GetFormulaSurface()),
            ExampleSteps.read(
                "analyze-workbook-findings",
                ExampleSelectors.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())),
        sourceAsset);
  }

  private static TaskStarterPlan customXmlStarter() {
    String taskId = "CUSTOM_XML_WORKFLOW";
    String mappingWorkbook = "custom-xml-assets/custom-xml-mapping.xlsx";
    String updateXml = "custom-xml-assets/custom-xml-update.xml";
    return TaskStarterPlanSupport.assetBackedStarter(
        taskId,
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.ExistingFile(mappingWorkbook),
            ExampleWorkbookPlans.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExampleSteps.read(
                "read-custom-xml-mappings",
                ExampleSelectors.workbook(),
                new WorkbookIntrospectionQuery.GetCustomXmlMappings()),
            ExampleSteps.read(
                "export-custom-xml-before-import",
                ExampleSelectors.workbook(),
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8")),
            ExampleSteps.step(
                "import-custom-xml",
                ExampleSelectors.workbook(),
                new StructuredMutationAction.ImportCustomXmlMapping(
                    new CustomXmlImportInput(
                        new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                        TextSourceInput.utf8File(updateXml)))),
            ExampleSteps.read(
                "read-imported-cells",
                ExampleSelectors.cells("Foglio1", "A1", "B1", "C1", "D1"),
                new SheetIntrospectionQuery.GetCells()),
            ExampleSteps.read(
                "export-custom-xml-after-import",
                ExampleSelectors.workbook(),
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8"))),
        mappingWorkbook,
        updateXml);
  }
}
