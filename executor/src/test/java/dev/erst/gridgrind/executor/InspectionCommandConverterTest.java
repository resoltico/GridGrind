package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelPivotTableSelection;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused coverage for step-native inspection-command conversion. */
class InspectionCommandConverterTest {
  @Test
  void convertsPackageSecurityInspectionStep() {
    InspectionStep step =
        new InspectionStep(
            "security", new WorkbookSelector.Current(), new InspectionQuery.GetPackageSecurity());
    WorkbookReadCommand.GetPackageSecurity command =
        assertInstanceOf(
            WorkbookReadCommand.GetPackageSecurity.class,
            InspectionCommandConverter.toReadCommand(step));

    assertEquals("security", command.stepId());
  }

  @Test
  void convertsChartInspectionSteps() {
    InspectionStep step =
        new InspectionStep(
            "charts", new ChartSelector.AllOnSheet("Budget"), new InspectionQuery.GetCharts());
    WorkbookReadCommand.GetCharts command =
        assertInstanceOf(
            WorkbookReadCommand.GetCharts.class, InspectionCommandConverter.toReadCommand(step));

    assertEquals("charts", command.stepId());
    assertEquals("Budget", command.sheetName());
  }

  @Test
  void convertsChartInspectionFromSheetSelector() {
    InspectionStep step =
        new InspectionStep(
            "charts-by-sheet", new SheetSelector.ByName("Budget"), new InspectionQuery.GetCharts());
    WorkbookReadCommand.GetCharts command =
        assertInstanceOf(
            WorkbookReadCommand.GetCharts.class, InspectionCommandConverter.toReadCommand(step));

    assertEquals("charts-by-sheet", command.stepId());
    assertEquals("Budget", command.sheetName());
  }

  @Test
  void convertsPivotInspectionSteps() {
    InspectionStep pivotStep =
        new InspectionStep(
            "pivots",
            new PivotTableSelector.ByNames(List.of("Sales Pivot 2026")),
            new InspectionQuery.GetPivotTables());
    InspectionStep pivotHealthStep =
        new InspectionStep(
            "pivot-health",
            new PivotTableSelector.All(),
            new InspectionQuery.AnalyzePivotTableHealth());

    WorkbookReadCommand.GetPivotTables pivotCommand =
        assertInstanceOf(
            WorkbookReadCommand.GetPivotTables.class,
            InspectionCommandConverter.toReadCommand(pivotStep));
    WorkbookReadCommand.AnalyzePivotTableHealth pivotHealthCommand =
        assertInstanceOf(
            WorkbookReadCommand.AnalyzePivotTableHealth.class,
            InspectionCommandConverter.toReadCommand(pivotHealthStep));

    assertEquals("pivots", pivotCommand.stepId());
    assertEquals(
        List.of("Sales Pivot 2026"),
        ((ExcelPivotTableSelection.ByNames) pivotCommand.selection()).names());
    assertInstanceOf(ExcelPivotTableSelection.All.class, pivotHealthCommand.selection());
  }

  @Test
  void convertsArrayFormulaInspectionFromSheetSelector() {
    InspectionStep step =
        new InspectionStep(
            "array-formulas",
            new SheetSelector.ByName("Calc"),
            new InspectionQuery.GetArrayFormulas());
    WorkbookReadCommand.GetArrayFormulas command =
        assertInstanceOf(
            WorkbookReadCommand.GetArrayFormulas.class,
            InspectionCommandConverter.toReadCommand(step));

    assertEquals("array-formulas", command.stepId());
    assertEquals(
        List.of("Calc"),
        ((dev.erst.gridgrind.excel.ExcelSheetSelection.Selected) command.selection()).sheetNames());
  }

  @Test
  void convertsWorkbookCustomXmlInspectionSteps() {
    InspectionStep mappingsStep =
        new InspectionStep(
            "custom-xml-mappings",
            new WorkbookSelector.Current(),
            new InspectionQuery.GetCustomXmlMappings());
    InspectionStep exportStep =
        new InspectionStep(
            "custom-xml-export",
            new WorkbookSelector.Current(),
            new InspectionQuery.ExportCustomXmlMapping(
                new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8"));

    WorkbookReadCommand.GetCustomXmlMappings mappingsCommand =
        assertInstanceOf(
            WorkbookReadCommand.GetCustomXmlMappings.class,
            InspectionCommandConverter.toReadCommand(mappingsStep));
    WorkbookReadCommand.ExportCustomXmlMapping exportCommand =
        assertInstanceOf(
            WorkbookReadCommand.ExportCustomXmlMapping.class,
            InspectionCommandConverter.toReadCommand(exportStep));

    assertEquals("custom-xml-mappings", mappingsCommand.stepId());
    assertEquals("custom-xml-export", exportCommand.stepId());
    assertEquals(1L, exportCommand.mapping().mapId());
    assertEquals("CORSO_mapping", exportCommand.mapping().name());
    assertEquals("UTF-8", exportCommand.encoding());
  }
}
