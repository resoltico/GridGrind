package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for Phase 7 source-backed plan resolution and failure semantics. */
class SourceBackedPlanResolverTest {
  @Test
  void sameOptionalReferenceDistinguishesPresenceMismatches() throws Exception {
    assertTrue(
        SourceBackedResolutionIdentitySupport.sameOptionalReference(
            Optional.empty(), Optional.empty()));
    assertFalse(
        SourceBackedResolutionIdentitySupport.sameOptionalReference(
            Optional.of("left"), Optional.empty()));
  }

  @Test
  void resolveRewritesTextFormulaAndBinarySourcesIntoInlineValues() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-resolve-");
    Path titleFile = workingDirectory.resolve("title.txt");
    Path formulaFile = workingDirectory.resolve("formula.txt");
    Path payloadFile = workingDirectory.resolve("payload.bin");
    Files.writeString(titleFile, "Quarterly Budget", StandardCharsets.UTF_8);
    Files.writeString(formulaFile, "=SUM(B2:B3)", StandardCharsets.UTF_8);
    byte[] payloadBytes = "payload".getBytes(StandardCharsets.UTF_8);
    Files.write(payloadFile, payloadBytes);

    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.Utf8File("title.txt")))),
                mutate(
                    new CellSelector.ByAddress("Budget", "A2"),
                    new MutationAction.SetCell(
                        new CellInput.Formula(new TextSourceInput.Utf8File("formula.txt")))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetShape(
                        new ShapeInput(
                            "QueueShape",
                            ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                            twoCellAnchor(),
                            "roundRect",
                            new TextSourceInput.StandardInput()))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetEmbeddedObject(
                        new EmbeddedObjectInput(
                            "OpsPayload",
                            "Ops payload",
                            "ops-payload.txt",
                            "open",
                            new BinarySourceInput.File("payload.bin"),
                            inlinePictureData(),
                            twoCellAnchor()))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSignatureLine(
                        new SignatureLineInput(
                            "OpsSignature",
                            twoCellAnchor(),
                            false,
                            java.util.Optional.of("Review before signing."),
                            java.util.Optional.of("Ada Lovelace"),
                            java.util.Optional.of("Finance"),
                            java.util.Optional.of("ada@example.com"),
                            java.util.Optional.empty(),
                            java.util.Optional.of("invalid"),
                            java.util.Optional.of(
                                new PictureDataInput(
                                    ExcelPictureFormat.PNG,
                                    new BinarySourceInput.File("payload.bin"))))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            new ExecutionInputBindings(workingDirectory, "Queue".getBytes(StandardCharsets.UTF_8)));

    MutationAction.SetCell titleAction =
        assertInstanceOf(
            MutationAction.SetCell.class, ((MutationStep) resolved.steps().get(0)).action());
    CellInput.Text titleValue = assertInstanceOf(CellInput.Text.class, titleAction.value());
    assertEquals("Quarterly Budget", ((TextSourceInput.Inline) titleValue.source()).text());

    MutationAction.SetCell formulaAction =
        assertInstanceOf(
            MutationAction.SetCell.class, ((MutationStep) resolved.steps().get(1)).action());
    CellInput.Formula formulaValue =
        assertInstanceOf(CellInput.Formula.class, formulaAction.value());
    assertEquals("SUM(B2:B3)", ((TextSourceInput.Inline) formulaValue.source()).text());

    MutationAction.SetShape shapeAction =
        assertInstanceOf(
            MutationAction.SetShape.class, ((MutationStep) resolved.steps().get(2)).action());
    assertEquals("Queue", ((TextSourceInput.Inline) shapeAction.shape().text()).text());

    MutationAction.SetEmbeddedObject embeddedObjectAction =
        assertInstanceOf(
            MutationAction.SetEmbeddedObject.class,
            ((MutationStep) resolved.steps().get(3)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(payloadBytes),
        ((BinarySourceInput.InlineBase64) embeddedObjectAction.embeddedObject().payload())
            .base64Data());

    MutationAction.SetSignatureLine signatureLineAction =
        assertInstanceOf(
            MutationAction.SetSignatureLine.class,
            ((MutationStep) resolved.steps().get(4)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(payloadBytes),
        ((BinarySourceInput.InlineBase64)
                signatureLineAction.signatureLine().plainSignature().orElseThrow().source())
            .base64Data());
  }

  @Test
  void requiresStandardInputTracksSourceBackedStandardInputFields() {
    WorkbookPlan standardInputPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.StandardInput()))),
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSignatureLine(
                        new SignatureLineInput(
                            "OpsSignature",
                            twoCellAnchor(),
                            false,
                            java.util.Optional.of("Review before signing."),
                            java.util.Optional.of("Ada Lovelace"),
                            java.util.Optional.of("Finance"),
                            java.util.Optional.of("ada@example.com"),
                            java.util.Optional.empty(),
                            java.util.Optional.of("invalid"),
                            java.util.Optional.of(
                                new PictureDataInput(
                                    ExcelPictureFormat.PNG,
                                    new BinarySourceInput.StandardInput())))))));
    WorkbookPlan fileBackedPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.Utf8File("title.txt"))))));

    assertTrue(SourceBackedPlanResolver.requiresStandardInput(standardInputPlan));
    assertFalse(SourceBackedPlanResolver.requiresStandardInput(fileBackedPlan));
  }

  @Test
  void resolveRewritesCustomXmlImportSourcesIntoInlineValues() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-custom-xml-resolve-");
    Files.writeString(
        workingDirectory.resolve("mapping.xml"),
        "<CORSO><NOME>Grid</NOME><DOCENTE>Grind</DOCENTE></CORSO>",
        StandardCharsets.UTF_8);

    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new WorkbookSelector.Current(),
                    new MutationAction.ImportCustomXmlMapping(
                        new CustomXmlImportInput(
                            new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                            new TextSourceInput.Utf8File("mapping.xml"))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(plan, new ExecutionInputBindings(workingDirectory));

    MutationAction.ImportCustomXmlMapping action =
        assertInstanceOf(
            MutationAction.ImportCustomXmlMapping.class,
            assertInstanceOf(MutationStep.class, resolved.steps().getFirst()).action());
    assertEquals(new CustomXmlMappingLocator(1L, "CORSO_mapping"), action.mapping().locator());
    assertEquals(
        "<CORSO><NOME>Grid</NOME><DOCENTE>Grind</DOCENTE></CORSO>",
        assertInstanceOf(TextSourceInput.Inline.class, action.mapping().xml()).text());
  }

  @Test
  void resolveRewritesSourceBackedTableKeySelectorsIntoInlineValues() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-selector-");
    Files.writeString(workingDirectory.resolve("item.txt"), "Hosting", StandardCharsets.UTF_8);

    TableCellSelector.ByColumnName selector =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(new TextSourceInput.Utf8File("item.txt"))),
            "Amount");
    WorkbookPlan plan =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                new InspectionStep("inspect-amount", selector, new InspectionQuery.GetCells()),
                new AssertionStep(
                    "assert-amount",
                    selector,
                    new Assertion.CellValue(new ExpectedCellValue.NumericValue(125.0)))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(plan, new ExecutionInputBindings(workingDirectory));

    InspectionStep resolvedInspection =
        assertInstanceOf(InspectionStep.class, resolved.steps().get(0));
    TableCellSelector.ByColumnName resolvedSelector =
        assertInstanceOf(TableCellSelector.ByColumnName.class, resolvedInspection.target());
    TableRowSelector.ByKeyCell resolvedRow =
        assertInstanceOf(TableRowSelector.ByKeyCell.class, resolvedSelector.row());
    CellInput.Text resolvedValue =
        assertInstanceOf(CellInput.Text.class, resolvedRow.expectedValue());
    assertEquals("Hosting", ((TextSourceInput.Inline) resolvedValue.source()).text());

    AssertionStep resolvedAssertion =
        assertInstanceOf(AssertionStep.class, resolved.steps().get(1));
    TableCellSelector.ByColumnName assertedSelector =
        assertInstanceOf(TableCellSelector.ByColumnName.class, resolvedAssertion.target());
    TableRowSelector.ByKeyCell assertedRow =
        assertInstanceOf(TableRowSelector.ByKeyCell.class, assertedSelector.row());
    CellInput.Text assertedValue =
        assertInstanceOf(CellInput.Text.class, assertedRow.expectedValue());
    assertEquals("Hosting", ((TextSourceInput.Inline) assertedValue.source()).text());
  }

  @Test
  void resolveRewritesChangedRangeRowsAndTextTitledCharts() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-rewrites-");
    Files.writeString(
        workingDirectory.resolve("range-cell.txt"), "Quarterly Budget", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("line-title.txt"), "Quarterly Trend", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("pie-title.txt"), "Category Mix", StandardCharsets.UTF_8);

    MutationStep rangeStep =
        new MutationStep(
            "set-range-file-backed",
            new RangeSelector.ByRange("Budget", "A1:B1"),
            new MutationAction.SetRange(
                List.of(
                    List.of(
                        new CellInput.Text(TextSourceInput.utf8File("range-cell.txt")),
                        new CellInput.Numeric(42.0d)))));
    MutationStep lineChartStep =
        new MutationStep(
            "set-line-chart-file-backed",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetChart(
                chartInput(
                    "Line",
                    twoCellAnchor(),
                    new ChartInput.Title.Text(TextSourceInput.utf8File("line-title.txt")),
                    null,
                    ExcelChartDisplayBlanksAs.ZERO,
                    true,
                    new ChartInput.Line(
                        false,
                        dev.erst.gridgrind.excel.foundation.ExcelChartGrouping.STANDARD,
                        List.of(chartSeries())))));
    MutationStep pieChartStep =
        new MutationStep(
            "set-pie-chart-file-backed",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetChart(
                chartInput(
                    "Pie",
                    twoCellAnchor(),
                    new ChartInput.Title.Text(TextSourceInput.utf8File("pie-title.txt")),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartInput.Pie(false, 15, List.of(chartSeries())))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(rangeStep, lineChartStep, pieChartStep)),
            new ExecutionInputBindings(workingDirectory));

    MutationStep resolvedRangeStep = assertInstanceOf(MutationStep.class, resolved.steps().get(0));
    assertNotSame(rangeStep, resolvedRangeStep);
    MutationAction.SetRange resolvedRange =
        assertInstanceOf(MutationAction.SetRange.class, resolvedRangeStep.action());
    CellInput.Text resolvedRangeCell =
        assertInstanceOf(CellInput.Text.class, resolvedRange.rows().getFirst().getFirst());
    assertEquals("Quarterly Budget", ((TextSourceInput.Inline) resolvedRangeCell.source()).text());

    MutationStep resolvedLineStep = assertInstanceOf(MutationStep.class, resolved.steps().get(1));
    assertNotSame(lineChartStep, resolvedLineStep);
    ChartInput resolvedLineChart =
        assertInstanceOf(MutationAction.SetChart.class, resolvedLineStep.action()).chart();
    assertInstanceOf(ChartInput.Line.class, resolvedLineChart.plots().getFirst());
    ChartInput.Title.Text lineTitle =
        assertInstanceOf(ChartInput.Title.Text.class, resolvedLineChart.title());
    assertEquals("Quarterly Trend", ((TextSourceInput.Inline) lineTitle.source()).text());

    MutationStep resolvedPieStep = assertInstanceOf(MutationStep.class, resolved.steps().get(2));
    assertNotSame(pieChartStep, resolvedPieStep);
    ChartInput resolvedPieChart =
        assertInstanceOf(MutationAction.SetChart.class, resolvedPieStep.action()).chart();
    assertInstanceOf(ChartInput.Pie.class, resolvedPieChart.plots().getFirst());
    ChartInput.Title.Text pieTitle =
        assertInstanceOf(ChartInput.Title.Text.class, resolvedPieChart.title());
    assertEquals("Category Mix", ((TextSourceInput.Inline) pieTitle.source()).text());
  }

  @Test
  void requiresStandardInputCoversAllMutationFamiliesAndNonMutationSteps() {
    MutationAction.SetRange fullyInlineRange =
        new MutationAction.SetRange(
            List.of(
                List.of(
                    new CellInput.Blank(),
                    new CellInput.RichText(
                        List.of(new RichTextRunInput(TextSourceInput.inline("Ada"), null))),
                    new CellInput.Numeric(1.0d),
                    new CellInput.BooleanValue(true),
                    new CellInput.Date(LocalDate.of(2026, 4, 18)),
                    new CellInput.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30)),
                    new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)")))));

    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(
                    new AssertionStep(
                        "assert-owner",
                        new CellSelector.ByAddress("Budget", "A1"),
                        new Assertion.CellValue(new ExpectedCellValue.Text("Owner")))))));
    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(
                    new InspectionStep(
                        "inspect-cells",
                        new CellSelector.ByAddress("Budget", "A1"),
                        new InspectionQuery.GetCells())))));

    assertFalse(
        SourceBackedPlanResolver.requiresStandardInput(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                mutations(
                    mutate(new RangeSelector.ByRange("Budget", "A1:G1"), fullyInlineRange)))));
    assertFalse(requiresStandardInputFor(new MutationAction.EnsureSheet()));
    assertFalse(requiresStandardInputFor(new MutationAction.SetCell(new CellInput.Numeric(1.0d))));
    assertFalse(
        requiresStandardInputFor(new MutationAction.SetCell(new CellInput.BooleanValue(true))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetCell(new CellInput.Date(LocalDate.of(2026, 4, 18)))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetCell(
                new CellInput.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30)))));

    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetCell(new CellInput.Text(TextSourceInput.standardInput()))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetRange(
                List.of(
                    List.of(
                        new CellInput.RichText(
                            List.of(
                                new RichTextRunInput(TextSourceInput.standardInput(), null))))))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.standardInput(),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty()))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty()))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.of(
                        List.of(new RichTextRunInput(TextSourceInput.standardInput(), null))),
                    java.util.Optional.empty()))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetPicture(
                new PictureInput(
                    "Logo",
                    inlinePictureData(),
                    twoCellAnchor(),
                    TextSourceInput.standardInput()))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetPicture(
                new PictureInput(
                    "Logo",
                    new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.standardInput()),
                    twoCellAnchor(),
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetChart(
                chartInput(
                    "Bar",
                    twoCellAnchor(),
                    new ChartInput.Title.Text(TextSourceInput.standardInput()),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartInput.Bar(
                        false,
                        dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection.COLUMN,
                        dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping.CLUSTERED,
                        null,
                        null,
                        List.of(chartSeries()))))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetChart(
                chartInput(
                    "Line",
                    twoCellAnchor(),
                    new ChartInput.Title.Formula("Budget!$A$1"),
                    null,
                    ExcelChartDisplayBlanksAs.ZERO,
                    true,
                    new ChartInput.Line(
                        false,
                        dev.erst.gridgrind.excel.foundation.ExcelChartGrouping.STANDARD,
                        List.of(chartSeries()))))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetChart(
                chartInput(
                    "Pie",
                    twoCellAnchor(),
                    new ChartInput.Title.None(),
                    null,
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    new ChartInput.Pie(false, 0, List.of(chartSeries()))))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetShape(
                new ShapeInput(
                    "Shape",
                    ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                    twoCellAnchor(),
                    "roundRect",
                    TextSourceInput.standardInput()))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetShape(
                new ShapeInput(
                    "Shape",
                    ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                    twoCellAnchor(),
                    "roundRect",
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.standardInput(),
                    inlinePictureData(),
                    twoCellAnchor()))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.inlineBase64("cGF5bG9hZA=="),
                    new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.standardInput()),
                    twoCellAnchor()))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.inlineBase64("cGF5bG9hZA=="),
                    inlinePictureData(),
                    twoCellAnchor()))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    new DataValidationPromptInput(
                        TextSourceInput.standardInput(), TextSourceInput.inline("Pick one"), false),
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    null,
                    new DataValidationErrorAlertInput(
                        ExcelDataValidationErrorStyle.STOP,
                        TextSourceInput.inline("Invalid"),
                        TextSourceInput.standardInput(),
                        false)))));
    assertFalse(
        requiresStandardInputFor(
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    null,
                    null))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetTable(
                new TableInput(
                    "BudgetTable",
                    "Budget",
                    "A1:B2",
                    false,
                    true,
                    new TableStyleInput.Named("TableStyleMedium2", false, false, true, false),
                    TextSourceInput.standardInput(),
                    false,
                    false,
                    false,
                    "",
                    "",
                    "",
                    List.of()))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.AppendRow(
                List.of(new CellInput.Formula(TextSourceInput.standardInput())))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextInput(
                        TextSourceInput.standardInput(),
                        TextSourceInput.inline(""),
                        TextSourceInput.inline(""))))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("left"),
                        TextSourceInput.standardInput(),
                        TextSourceInput.inline("right"))))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetPrintLayout(
                printLayoutWithHeader(
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("left"),
                        TextSourceInput.inline("center"),
                        TextSourceInput.standardInput())))));
    assertTrue(
        requiresStandardInputFor(
            new MutationAction.SetPrintLayout(
                printLayoutWithFooter(
                    new HeaderFooterTextInput(
                        TextSourceInput.standardInput(),
                        TextSourceInput.inline(""),
                        TextSourceInput.inline(""))))));
  }

  @Test
  void resolveCoversRemainingMutationFamiliesAndPreservesNonMutationSteps() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-phase7-");
    Files.writeString(
        workingDirectory.resolve("comment.txt"), "Ada Lovelace", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("run1.txt"), "Ada", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("run2.txt"), " Lovelace", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("picture-description.txt"),
        "Queue preview",
        StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("chart-title.txt"), "Quarterly Trend", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("shape.txt"), "Queue shape", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("prompt-title.txt"), "Choose", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("prompt-text.txt"), "Pick one value", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("error-title.txt"), "Invalid", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("error-text.txt"), "Try again", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("table-comment.txt"),
        "Quarterly budget table",
        StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("header-left.txt"), "Left", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("header-center.txt"), "Center", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("header-right.txt"), "Right", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("formula.txt"), "=SUM(B2:B3)", StandardCharsets.UTF_8);
    Files.write(
        workingDirectory.resolve("payload.bin"), "payload".getBytes(StandardCharsets.UTF_8));
    Files.write(workingDirectory.resolve("preview.bin"), pngBytes());

    AssertionStep assertion =
        new AssertionStep(
            "assert-budget",
            new CellSelector.ByAddress("Budget", "A1"),
            new Assertion.CellValue(new ExpectedCellValue.Text("Quarterly Budget")));
    InspectionStep inspection =
        new InspectionStep(
            "inspect-budget",
            new CellSelector.ByAddress("Budget", "A1"),
            new InspectionQuery.GetCells());

    WorkbookPlan plan =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                new MutationStep(
                    "step-01-set-comment",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetComment(
                        new CommentInput(
                            TextSourceInput.utf8File("comment.txt"),
                            "Ada",
                            false,
                            java.util.Optional.of(
                                List.of(
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run1.txt"), null),
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run2.txt"), null))),
                            java.util.Optional.empty()))),
                new MutationStep(
                    "step-02-set-picture",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetPicture(
                        new PictureInput(
                            "Logo",
                            new PictureDataInput(
                                ExcelPictureFormat.PNG, BinarySourceInput.file("preview.bin")),
                            twoCellAnchor(),
                            TextSourceInput.utf8File("picture-description.txt")))),
                new MutationStep(
                    "step-03-set-bar-chart",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetChart(
                        chartInput(
                            "Bar",
                            twoCellAnchor(),
                            new ChartInput.Title.Text(TextSourceInput.utf8File("chart-title.txt")),
                            null,
                            ExcelChartDisplayBlanksAs.GAP,
                            true,
                            new ChartInput.Bar(
                                false,
                                dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection.COLUMN,
                                dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping.CLUSTERED,
                                null,
                                null,
                                List.of(chartSeries()))))),
                new MutationStep(
                    "step-04-set-line-chart",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetChart(
                        chartInput(
                            "Line",
                            twoCellAnchor(),
                            new ChartInput.Title.Formula("Budget!$A$1"),
                            null,
                            ExcelChartDisplayBlanksAs.ZERO,
                            true,
                            new ChartInput.Line(
                                false,
                                dev.erst.gridgrind.excel.foundation.ExcelChartGrouping.STANDARD,
                                List.of(chartSeries()))))),
                new MutationStep(
                    "step-05-set-pie-chart",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetChart(
                        chartInput(
                            "Pie",
                            twoCellAnchor(),
                            new ChartInput.Title.None(),
                            null,
                            ExcelChartDisplayBlanksAs.GAP,
                            true,
                            new ChartInput.Pie(false, 0, List.of(chartSeries()))))),
                new MutationStep(
                    "step-06-set-shape",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetShape(
                        new ShapeInput(
                            "Shape",
                            ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                            twoCellAnchor(),
                            "roundRect",
                            TextSourceInput.utf8File("shape.txt")))),
                new MutationStep(
                    "step-07-set-embedded",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetEmbeddedObject(
                        new EmbeddedObjectInput(
                            "Payload",
                            "Payload",
                            "payload.bin",
                            "open",
                            BinarySourceInput.file("payload.bin"),
                            new PictureDataInput(
                                ExcelPictureFormat.PNG, BinarySourceInput.file("preview.bin")),
                            twoCellAnchor()))),
                new MutationStep(
                    "step-08-set-validation",
                    new RangeSelector.ByRange("Budget", "A1"),
                    new MutationAction.SetDataValidation(
                        new DataValidationInput(
                            new DataValidationRuleInput.WholeNumber(
                                ExcelComparisonOperator.BETWEEN, "1", "10"),
                            false,
                            false,
                            new DataValidationPromptInput(
                                TextSourceInput.utf8File("prompt-title.txt"),
                                TextSourceInput.utf8File("prompt-text.txt"),
                                false),
                            new DataValidationErrorAlertInput(
                                ExcelDataValidationErrorStyle.STOP,
                                TextSourceInput.utf8File("error-title.txt"),
                                TextSourceInput.utf8File("error-text.txt"),
                                false)))),
                new MutationStep(
                    "step-09-set-table",
                    new dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet(
                        "BudgetTable", "Budget"),
                    new MutationAction.SetTable(
                        new TableInput(
                            "BudgetTable",
                            "Budget",
                            "A1:B2",
                            false,
                            true,
                            new TableStyleInput.Named(
                                "TableStyleMedium2", false, false, true, false),
                            TextSourceInput.utf8File("table-comment.txt"),
                            false,
                            false,
                            false,
                            "",
                            "",
                            "",
                            List.of()))),
                new MutationStep(
                    "step-10-append-row",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.AppendRow(
                        List.of(
                            new CellInput.Text(TextSourceInput.utf8File("chart-title.txt")),
                            new CellInput.RichText(
                                List.of(
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run1.txt"), null),
                                    new RichTextRunInput(
                                        TextSourceInput.utf8File("run2.txt"), null))),
                            new CellInput.Formula(TextSourceInput.utf8File("formula.txt"))))),
                new MutationStep(
                    "step-11-print-layout",
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetPrintLayout(
                        printLayoutWithHeader(
                            new HeaderFooterTextInput(
                                TextSourceInput.utf8File("header-left.txt"),
                                TextSourceInput.utf8File("header-center.txt"),
                                TextSourceInput.utf8File("header-right.txt"))))),
                assertion,
                inspection));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            new ExecutionInputBindings(
                workingDirectory, "stdin-text".getBytes(StandardCharsets.UTF_8)));

    MutationAction.SetComment commentAction =
        assertInstanceOf(
            MutationAction.SetComment.class, ((MutationStep) resolved.steps().get(0)).action());
    assertEquals("Ada Lovelace", ((TextSourceInput.Inline) commentAction.comment().text()).text());
    assertEquals(
        "Ada",
        ((TextSourceInput.Inline) commentAction.comment().runs().orElseThrow().getFirst().source())
            .text());

    MutationAction.SetPicture pictureAction =
        assertInstanceOf(
            MutationAction.SetPicture.class, ((MutationStep) resolved.steps().get(1)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(pngBytes()),
        ((BinarySourceInput.InlineBase64) pictureAction.picture().image().source()).base64Data());
    assertEquals(
        "Queue preview", ((TextSourceInput.Inline) pictureAction.picture().description()).text());

    MutationAction.SetChart barChart =
        assertInstanceOf(
            MutationAction.SetChart.class, ((MutationStep) resolved.steps().get(2)).action());
    ChartInput resolvedBarChart = barChart.chart();
    assertInstanceOf(ChartInput.Bar.class, resolvedBarChart.plots().getFirst());
    assertEquals(
        "Quarterly Trend",
        ((TextSourceInput.Inline)
                assertInstanceOf(ChartInput.Title.Text.class, resolvedBarChart.title()).source())
            .text());

    MutationAction.SetChart lineChart =
        assertInstanceOf(
            MutationAction.SetChart.class, ((MutationStep) resolved.steps().get(3)).action());
    assertInstanceOf(ChartInput.Line.class, lineChart.chart().plots().getFirst());
    assertInstanceOf(ChartInput.Title.Formula.class, lineChart.chart().title());

    MutationAction.SetChart pieChart =
        assertInstanceOf(
            MutationAction.SetChart.class, ((MutationStep) resolved.steps().get(4)).action());
    assertInstanceOf(ChartInput.Pie.class, pieChart.chart().plots().getFirst());
    assertInstanceOf(ChartInput.Title.None.class, pieChart.chart().title());

    MutationAction.SetShape shapeAction =
        assertInstanceOf(
            MutationAction.SetShape.class, ((MutationStep) resolved.steps().get(5)).action());
    assertEquals("Queue shape", ((TextSourceInput.Inline) shapeAction.shape().text()).text());

    MutationAction.SetEmbeddedObject embeddedAction =
        assertInstanceOf(
            MutationAction.SetEmbeddedObject.class,
            ((MutationStep) resolved.steps().get(6)).action());
    assertEquals(
        Base64.getEncoder().encodeToString("payload".getBytes(StandardCharsets.UTF_8)),
        ((BinarySourceInput.InlineBase64) embeddedAction.embeddedObject().payload()).base64Data());

    MutationAction.SetDataValidation validationAction =
        assertInstanceOf(
            MutationAction.SetDataValidation.class,
            ((MutationStep) resolved.steps().get(7)).action());
    assertEquals(
        "Choose", ((TextSourceInput.Inline) validationAction.validation().prompt().title()).text());
    assertEquals(
        "Try again",
        ((TextSourceInput.Inline) validationAction.validation().errorAlert().text()).text());

    MutationAction.SetTable tableAction =
        assertInstanceOf(
            MutationAction.SetTable.class, ((MutationStep) resolved.steps().get(8)).action());
    assertEquals(
        "Quarterly budget table", ((TextSourceInput.Inline) tableAction.table().comment()).text());

    MutationAction.AppendRow appendRow =
        assertInstanceOf(
            MutationAction.AppendRow.class, ((MutationStep) resolved.steps().get(9)).action());
    assertEquals(
        "SUM(B2:B3)",
        ((TextSourceInput.Inline)
                assertInstanceOf(CellInput.Formula.class, appendRow.values().get(2)).source())
            .text());

    MutationAction.SetPrintLayout printLayout =
        assertInstanceOf(
            MutationAction.SetPrintLayout.class,
            ((MutationStep) resolved.steps().get(10)).action());
    assertEquals(
        "Left", ((TextSourceInput.Inline) printLayout.printLayout().header().left()).text());
    assertEquals(
        "Center", ((TextSourceInput.Inline) printLayout.printLayout().header().center()).text());
    assertEquals(
        "Right", ((TextSourceInput.Inline) printLayout.printLayout().header().right()).text());

    assertSame(assertion, resolved.steps().get(11));
    assertSame(inspection, resolved.steps().get(12));
  }

  @Test
  void resolveRejectsWhitespaceDirectoryInvalidPathsAndEmptyBinarySources() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-errors-");
    Files.writeString(workingDirectory.resolve("blank.txt"), "   ", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("empty.txt"), "", StandardCharsets.UTF_8);
    Path directory = Files.createDirectory(workingDirectory.resolve("dir"));
    Path textLoop = workingDirectory.resolve("text-loop.txt");
    Path binaryLoop = workingDirectory.resolve("binary-loop.bin");
    Files.createSymbolicLink(textLoop, Path.of("text-loop.txt"));
    Files.createSymbolicLink(binaryLoop, Path.of("binary-loop.bin"));

    WorkbookPlan blankCellTextPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(TextSourceInput.utf8File("blank.txt"))))));
    assertEquals(
        "cell text must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SourceBackedPlanResolver.resolve(
                        blankCellTextPlan, new ExecutionInputBindings(workingDirectory)))
            .getMessage());

    WorkbookPlan emptyRichTextRunPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.RichText(
                            List.of(
                                new RichTextRunInput(
                                    TextSourceInput.utf8File("empty.txt"), null)))))));
    assertEquals(
        "rich-text run must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SourceBackedPlanResolver.resolve(
                        emptyRichTextRunPlan, new ExecutionInputBindings(workingDirectory)))
            .getMessage());

    WorkbookPlan directoryPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(
                            TextSourceInput.utf8File(directory.getFileName().toString()))))));
    InputSourceReadException directoryFailure =
        assertThrows(
            InputSourceReadException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    directoryPlan, new ExecutionInputBindings(workingDirectory)));
    assertTrue(directoryFailure.getMessage().contains("must resolve to a file"));

    WorkbookPlan invalidPathPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(TextSourceInput.utf8File("\0bad"))))));
    InputSourceReadException invalidPathFailure =
        assertThrows(
            InputSourceReadException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    invalidPathPlan, new ExecutionInputBindings(workingDirectory)));
    assertTrue(invalidPathFailure.getMessage().contains("Invalid cell text path"));

    WorkbookPlan textLoopPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(TextSourceInput.utf8File("text-loop.txt"))))));
    InputSourceReadException textLoopFailure =
        assertThrows(
            InputSourceReadException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    textLoopPlan, new ExecutionInputBindings(workingDirectory)));
    assertTrue(textLoopFailure.getMessage().contains("Failed to read cell text file"));

    WorkbookPlan missingBinaryPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetPicture(
                        new PictureInput(
                            "Logo",
                            new PictureDataInput(
                                ExcelPictureFormat.PNG, BinarySourceInput.file("missing.bin")),
                            twoCellAnchor(),
                            null)))));
    InputSourceNotFoundException missingBinaryFailure =
        assertThrows(
            InputSourceNotFoundException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    missingBinaryPlan, new ExecutionInputBindings(workingDirectory)));
    assertEquals("picture payload", missingBinaryFailure.inputKind());

    WorkbookPlan binaryLoopPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetPicture(
                        new PictureInput(
                            "Logo",
                            new PictureDataInput(
                                ExcelPictureFormat.PNG, BinarySourceInput.file("binary-loop.bin")),
                            twoCellAnchor(),
                            null)))));
    InputSourceReadException binaryLoopFailure =
        assertThrows(
            InputSourceReadException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    binaryLoopPlan, new ExecutionInputBindings(workingDirectory)));
    assertTrue(binaryLoopFailure.getMessage().contains("Failed to read picture payload file"));

    WorkbookPlan emptyBinaryPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetEmbeddedObject(
                        new EmbeddedObjectInput(
                            "Payload",
                            "Payload",
                            "payload.bin",
                            "open",
                            BinarySourceInput.standardInput(),
                            inlinePictureData(),
                            twoCellAnchor())))));
    assertEquals(
        "embedded-object payload must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SourceBackedPlanResolver.resolve(
                        emptyBinaryPlan, new ExecutionInputBindings(workingDirectory, new byte[0])))
            .getMessage());
  }

  @Test
  void helperPredicatesCoverRemainingSourceBackedBranches() throws IOException {
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new AssertionStep(
                "assert-owner",
                new CellSelector.ByAddress("Budget", "A1"),
                new Assertion.CellValue(new ExpectedCellValue.Text("Owner")))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new InspectionStep(
                "inspect-cells",
                new CellSelector.ByAddress("Budget", "A1"),
                new InspectionQuery.GetCells())));

    assertFalse(SourceBackedInputRequirements.requiresStandardInput(new CellInput.Numeric(1.0d)));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(new CellInput.BooleanValue(true)));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellInput.Date(LocalDate.of(2026, 4, 18))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new CellInput.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30))));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new CommentInput(
                TextSourceInput.inline("Ada"),
                "Ada",
                false,
                java.util.Optional.of(
                    List.of(new RichTextRunInput(TextSourceInput.inline("Ada"), null))),
                java.util.Optional.empty())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CommentInput(
                TextSourceInput.inline("Ada"),
                "Ada",
                false,
                java.util.Optional.of(
                    List.of(new RichTextRunInput(TextSourceInput.standardInput(), null))),
                java.util.Optional.empty())));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new PictureInput(
                "Logo", inlinePictureData(), twoCellAnchor(), TextSourceInput.inline("Logo"))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new PictureInput(
                "Logo",
                new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.standardInput()),
                twoCellAnchor(),
                null)));

    assertFalse(SourceBackedInputRequirements.requiresStandardInput(new ChartInput.Title.None()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartInput.Title.Formula("Budget!$A$1")));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartInput.Title.Text(TextSourceInput.inline("Title"))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartInput.Title.Text(TextSourceInput.standardInput())));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ShapeInput(
                "Shape",
                ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                twoCellAnchor(),
                "roundRect",
                TextSourceInput.inline("Shape"))));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ShapeInput(
                "Shape",
                ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                twoCellAnchor(),
                "roundRect",
                null)));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new DataValidationInput(
                new DataValidationRuleInput.ExplicitList(List.of("Open")),
                false,
                false,
                new DataValidationPromptInput(
                    TextSourceInput.inline("Prompt"), TextSourceInput.inline("Body"), false),
                new DataValidationErrorAlertInput(
                    ExcelDataValidationErrorStyle.STOP,
                    TextSourceInput.inline("Alert"),
                    TextSourceInput.inline("Body"),
                    false))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DataValidationInput(
                new DataValidationRuleInput.ExplicitList(List.of("Open")),
                false,
                false,
                null,
                new DataValidationErrorAlertInput(
                    ExcelDataValidationErrorStyle.STOP,
                    TextSourceInput.standardInput(),
                    TextSourceInput.inline("Body"),
                    false))));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new DataValidationPromptInput(
                TextSourceInput.inline("Prompt"), TextSourceInput.inline("Body"), false)));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DataValidationPromptInput(
                TextSourceInput.inline("Prompt"), TextSourceInput.standardInput(), false)));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceInput.inline("Alert"),
                TextSourceInput.inline("Body"),
                false)));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceInput.standardInput(),
                TextSourceInput.inline("Body"),
                false)));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            printLayoutWithHeader(
                new HeaderFooterTextInput(
                    TextSourceInput.inline("left"),
                    TextSourceInput.inline("center"),
                    TextSourceInput.inline("right")))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            printLayoutWithFooter(
                new HeaderFooterTextInput(
                    TextSourceInput.standardInput(),
                    TextSourceInput.inline(""),
                    TextSourceInput.inline("")))));

    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new HeaderFooterTextInput(
                TextSourceInput.inline("left"),
                TextSourceInput.inline("center"),
                TextSourceInput.inline("right"))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new HeaderFooterTextInput(
                TextSourceInput.inline("left"),
                TextSourceInput.inline("center"),
                TextSourceInput.standardInput())));

    Path absoluteFile = Files.createTempFile("gridgrind-source-backed-absolute-", ".txt");
    assertEquals(
        absoluteFile.toAbsolutePath().normalize(),
        SourceBackedPathResolver.resolvePath(absoluteFile.toString(), Path.of(""), "cell text"));
  }

  @Test
  void resolveReportsMissingFilesAndUnavailableStandardInput() throws IOException {
    WorkbookPlan missingFilePlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.Utf8File("missing.txt"))))));
    InputSourceNotFoundException notFound =
        assertThrows(
            InputSourceNotFoundException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    missingFilePlan, new ExecutionInputBindings(Path.of(""))));
    assertEquals("cell text", notFound.inputKind());

    WorkbookPlan standardInputPlan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.StandardInput())))));
    InputSourceUnavailableException unavailable =
        assertThrows(
            InputSourceUnavailableException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    standardInputPlan, new ExecutionInputBindings(Path.of(""))));
    assertEquals("cell text", unavailable.inputKind());
  }

  @Test
  void resolvePreservesWhitespaceOnlyRichTextRuns() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-rich-text-");
    Path whitespaceFile = workingDirectory.resolve("space.txt");
    Files.writeString(whitespaceFile, " ", StandardCharsets.UTF_8);

    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(
                        new CellInput.RichText(
                            List.of(
                                new RichTextRunInput(
                                    new TextSourceInput.Utf8File("space.txt"), null)))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(plan, new ExecutionInputBindings(workingDirectory));

    MutationAction.SetCell richTextAction =
        assertInstanceOf(
            MutationAction.SetCell.class, ((MutationStep) resolved.steps().getFirst()).action());
    CellInput.RichText richTextValue =
        assertInstanceOf(CellInput.RichText.class, richTextAction.value());
    assertEquals(" ", ((TextSourceInput.Inline) richTextValue.runs().getFirst().source()).text());
  }

  @Test
  void requiresStandardInputTracksSemanticTableSelectors() {
    TableCellSelector.ByColumnName standardInputTarget =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(new TextSourceInput.StandardInput())),
            "Amount");
    TableCellSelector.ByColumnName inlineTarget =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(TextSourceInput.inline("Hosting"))),
            "Amount");

    WorkbookPlan standardInputPlan =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                new InspectionStep(
                    "inspect-amount", standardInputTarget, new InspectionQuery.GetCells())));
    WorkbookPlan inlinePlan =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                new InspectionStep(
                    "inspect-amount", inlineTarget, new InspectionQuery.GetCells())));

    assertTrue(SourceBackedPlanResolver.requiresStandardInput(standardInputPlan));
    assertFalse(SourceBackedPlanResolver.requiresStandardInput(inlinePlan));
  }

  @Test
  void resolvePreservesCanonicalStepInstancesWhenNothingChanges() throws IOException {
    MutationStep mutationStep =
        new MutationStep(
            "mutate-inline",
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetCell(new CellInput.Text(TextSourceInput.inline("Owner"))));
    InspectionStep inspectionStep =
        new InspectionStep(
            "inspect-inline",
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByName("BudgetTable"),
                    "Item",
                    new CellInput.Text(TextSourceInput.inline("Hosting"))),
                "Amount"),
            new InspectionQuery.GetCells());
    AssertionStep assertionStep =
        new AssertionStep(
            "assert-inline",
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByName("BudgetTable"),
                    "Item",
                    new CellInput.Text(TextSourceInput.inline("Hosting"))),
                "Amount"),
            new Assertion.CellValue(new ExpectedCellValue.NumericValue(125.0)));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(mutationStep, inspectionStep, assertionStep)),
            new ExecutionInputBindings(Path.of("")));

    assertSame(mutationStep, resolved.steps().get(0));
    assertSame(inspectionStep, resolved.steps().get(1));
    assertSame(assertionStep, resolved.steps().get(2));
  }

  @Test
  void resolvePreservesAlreadyInlineMutationFamiliesWhenNothingNeedsInlining() throws IOException {
    MutationStep rangeStep =
        new MutationStep(
            "range-inline",
            new RangeSelector.ByRange("Budget", "A1:B1"),
            new MutationAction.SetRange(
                List.of(
                    List.of(
                        new CellInput.Text(TextSourceInput.inline("Owner")),
                        new CellInput.Numeric(42.0d)))));
    MutationStep commentStep =
        new MutationStep(
            "comment-inline",
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty())));
    MutationStep pictureStep =
        new MutationStep(
            "picture-inline",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPicture(
                new PictureInput(
                    "Logo", inlinePictureData(), twoCellAnchor(), TextSourceInput.inline("Logo"))));
    MutationStep shapeStep =
        new MutationStep(
            "shape-inline",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetShape(
                new ShapeInput(
                    "Shape",
                    ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                    twoCellAnchor(),
                    "roundRect",
                    TextSourceInput.inline("Shape"))));
    MutationStep validationStep =
        new MutationStep(
            "validation-inline",
            new RangeSelector.ByRange("Budget", "A1"),
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    new DataValidationPromptInput(
                        TextSourceInput.inline("Prompt"), TextSourceInput.inline("Body"), false),
                    new DataValidationErrorAlertInput(
                        ExcelDataValidationErrorStyle.STOP,
                        TextSourceInput.inline("Error"),
                        TextSourceInput.inline("Try again"),
                        false))));
    MutationStep tableStep =
        new MutationStep(
            "table-inline",
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
            new MutationAction.SetTable(
                new TableInput(
                    "BudgetTable",
                    "Budget",
                    "A1:B2",
                    false,
                    true,
                    new TableStyleInput.Named("TableStyleMedium2", false, false, true, false),
                    TextSourceInput.inline("Budget table"),
                    false,
                    false,
                    false,
                    "",
                    "",
                    "",
                    List.of())));
    MutationStep printLayoutStep =
        new MutationStep(
            "print-layout-inline",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPrintLayout(
                printLayoutWithHeaderAndFooter(
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("left"),
                        TextSourceInput.inline("center"),
                        TextSourceInput.inline("right")),
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("footer-left"),
                        TextSourceInput.inline("footer-center"),
                        TextSourceInput.inline("footer-right")))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(
                    rangeStep,
                    commentStep,
                    pictureStep,
                    shapeStep,
                    validationStep,
                    tableStep,
                    printLayoutStep)),
            new ExecutionInputBindings(Path.of("")));

    assertSame(rangeStep, resolved.steps().get(0));
    assertSame(commentStep, resolved.steps().get(1));
    assertSame(pictureStep, resolved.steps().get(2));
    assertSame(shapeStep, resolved.steps().get(3));
    assertSame(validationStep, resolved.steps().get(4));
    assertSame(tableStep, resolved.steps().get(5));
    assertSame(printLayoutStep, resolved.steps().get(6));
  }

  @Test
  void resolveHandlesMixedInlineAndSourceBackedFieldsWithoutLosingCanonicalStructure()
      throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-mixed-");
    Files.writeString(
        workingDirectory.resolve("comment-run.txt"), " Lovelace", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("picture-description.txt"),
        "Queue preview",
        StandardCharsets.UTF_8);
    Files.write(workingDirectory.resolve("preview.bin"), pngBytes());
    Files.writeString(
        workingDirectory.resolve("prompt-body.txt"), "Pick one value", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("error-body.txt"), "Try again", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("header-center.txt"), "Centered", StandardCharsets.UTF_8);

    MutationStep formulaStep =
        new MutationStep(
            "formula-inline-prefix",
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("=SUM(B2:B3)"))));
    MutationStep commentStep =
        new MutationStep(
            "comment-mixed",
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada Lovelace"),
                    "Ada",
                    false,
                    java.util.Optional.of(
                        List.of(
                            new RichTextRunInput(TextSourceInput.inline("Ada"), null),
                            new RichTextRunInput(
                                TextSourceInput.utf8File("comment-run.txt"), null))),
                    java.util.Optional.empty())));
    MutationStep pictureStep =
        new MutationStep(
            "picture-mixed",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPicture(
                new PictureInput(
                    "Logo",
                    new PictureDataInput(
                        ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64("cGF5bG9hZA")),
                    twoCellAnchor(),
                    TextSourceInput.utf8File("picture-description.txt"))));
    MutationStep embeddedObjectStep =
        new MutationStep(
            "embedded-mixed",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "Payload",
                    "Payload",
                    "payload.bin",
                    "open",
                    BinarySourceInput.inlineBase64("cGF5bG9hZA=="),
                    new PictureDataInput(
                        ExcelPictureFormat.PNG, BinarySourceInput.file("preview.bin")),
                    twoCellAnchor())));
    MutationStep promptValidationStep =
        new MutationStep(
            "validation-prompt-mixed",
            new RangeSelector.ByRange("Budget", "A1"),
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    new DataValidationPromptInput(
                        TextSourceInput.inline("Prompt"),
                        TextSourceInput.utf8File("prompt-body.txt"),
                        false),
                    null)));
    MutationStep errorValidationStep =
        new MutationStep(
            "validation-error-mixed",
            new RangeSelector.ByRange("Budget", "B1"),
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.ExplicitList(List.of("Open")),
                    false,
                    false,
                    null,
                    new DataValidationErrorAlertInput(
                        ExcelDataValidationErrorStyle.STOP,
                        TextSourceInput.inline("Error"),
                        TextSourceInput.utf8File("error-body.txt"),
                        false))));
    MutationStep printLayoutStep =
        new MutationStep(
            "print-layout-mixed",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPrintLayout(
                printLayoutWithHeaderAndFooter(
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("left"),
                        TextSourceInput.utf8File("header-center.txt"),
                        TextSourceInput.inline("right")),
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("footer-left"),
                        TextSourceInput.inline("footer-center"),
                        TextSourceInput.inline("footer-right")))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(
                    formulaStep,
                    commentStep,
                    pictureStep,
                    embeddedObjectStep,
                    promptValidationStep,
                    errorValidationStep,
                    printLayoutStep)),
            new ExecutionInputBindings(workingDirectory));

    MutationAction.SetCell resolvedFormulaAction =
        assertInstanceOf(
            MutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(0)).action());
    CellInput.Formula resolvedFormula =
        assertInstanceOf(CellInput.Formula.class, resolvedFormulaAction.value());
    assertEquals("SUM(B2:B3)", ((TextSourceInput.Inline) resolvedFormula.source()).text());

    MutationAction.SetComment resolvedCommentAction =
        assertInstanceOf(
            MutationAction.SetComment.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(1)).action());
    assertEquals(
        " Lovelace",
        ((TextSourceInput.Inline)
                resolvedCommentAction.comment().runs().orElseThrow().get(1).source())
            .text());

    MutationAction.SetPicture resolvedPictureAction =
        assertInstanceOf(
            MutationAction.SetPicture.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(2)).action());
    assertEquals(
        "cGF5bG9hZA==",
        ((BinarySourceInput.InlineBase64) resolvedPictureAction.picture().image().source())
            .base64Data());
    assertEquals(
        "Queue preview",
        ((TextSourceInput.Inline) resolvedPictureAction.picture().description()).text());

    MutationAction.SetEmbeddedObject resolvedEmbeddedAction =
        assertInstanceOf(
            MutationAction.SetEmbeddedObject.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(3)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(pngBytes()),
        ((BinarySourceInput.InlineBase64)
                resolvedEmbeddedAction.embeddedObject().previewImage().source())
            .base64Data());

    MutationAction.SetDataValidation resolvedPromptValidation =
        assertInstanceOf(
            MutationAction.SetDataValidation.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(4)).action());
    assertEquals(
        "Pick one value",
        ((TextSourceInput.Inline) resolvedPromptValidation.validation().prompt().text()).text());

    MutationAction.SetDataValidation resolvedErrorValidation =
        assertInstanceOf(
            MutationAction.SetDataValidation.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(5)).action());
    assertEquals(
        "Try again",
        ((TextSourceInput.Inline) resolvedErrorValidation.validation().errorAlert().text()).text());

    MutationAction.SetPrintLayout resolvedPrintLayout =
        assertInstanceOf(
            MutationAction.SetPrintLayout.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(6)).action());
    assertEquals(
        "Centered",
        ((TextSourceInput.Inline) resolvedPrintLayout.printLayout().header().center()).text());
  }

  @Test
  void resolveCoversMixedTargetAndActionIdentityBranches() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-identity-branches-");
    Files.writeString(workingDirectory.resolve("item.txt"), "Hosting", StandardCharsets.UTF_8);
    Files.writeString(workingDirectory.resolve("value.txt"), "Ada", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("footer-right.txt"), "right-file", StandardCharsets.UTF_8);

    MutationStep targetAndActionChanged =
        new MutationStep(
            "target-and-action-changed",
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByName("BudgetTable"),
                    "Item",
                    new CellInput.Text(TextSourceInput.utf8File("item.txt"))),
                "Amount"),
            new MutationAction.SetCell(new CellInput.Text(TextSourceInput.utf8File("value.txt"))));
    MutationStep pictureDescriptionOnlyChanged =
        new MutationStep(
            "picture-description-only",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPicture(
                new PictureInput(
                    "Logo",
                    inlinePictureData(),
                    twoCellAnchor(),
                    TextSourceInput.utf8File("value.txt"))));
    MutationStep footerOnlyChanged =
        new MutationStep(
            "footer-only-changed",
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPrintLayout(
                printLayoutWithHeaderAndFooter(
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("header-left"),
                        TextSourceInput.inline("header-center"),
                        TextSourceInput.inline("header-right")),
                    new HeaderFooterTextInput(
                        TextSourceInput.inline("footer-left"),
                        TextSourceInput.inline("footer-center"),
                        TextSourceInput.utf8File("footer-right.txt")))));
    MutationStep formulaAlreadyInline =
        new MutationStep(
            "formula-inline-stable",
            new CellSelector.ByAddress("Budget", "B2"),
            new MutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)"))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(
                    targetAndActionChanged,
                    pictureDescriptionOnlyChanged,
                    footerOnlyChanged,
                    formulaAlreadyInline)),
            new ExecutionInputBindings(workingDirectory));

    assertNotSame(targetAndActionChanged, resolved.steps().get(0));
    assertNotSame(pictureDescriptionOnlyChanged, resolved.steps().get(1));
    assertNotSame(footerOnlyChanged, resolved.steps().get(2));
    assertSame(formulaAlreadyInline, resolved.steps().get(3));
  }

  @Test
  void resolveCanonicalizesFormulaSourcesAcrossStableChangedAndFileBackedVariants()
      throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-formula-source-");
    Files.writeString(
        workingDirectory.resolve("formula.txt"), "=SUM(C1:C2)", StandardCharsets.UTF_8);
    ExecutionInputBindings bindings = new ExecutionInputBindings(workingDirectory);

    TextSourceInput.Inline stableSource = TextSourceInput.inline("SUM(A1:A2)");
    assertSame(stableSource, SourceBackedPlanResolver.resolveFormulaSource(stableSource, bindings));

    TextSourceInput.Inline prefixedSource = TextSourceInput.inline("=SUM(B1:B2)");
    assertEquals(
        "SUM(B1:B2)",
        ((TextSourceInput.Inline)
                SourceBackedPlanResolver.resolveFormulaSource(prefixedSource, bindings))
            .text());

    assertEquals(
        "SUM(C1:C2)",
        ((TextSourceInput.Inline)
                SourceBackedPlanResolver.resolveFormulaSource(
                    TextSourceInput.utf8File("formula.txt"), bindings))
            .text());

    MutationStep stableInline =
        new MutationStep(
            "formula-inline-stable",
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)"))));
    MutationStep prefixedInline =
        new MutationStep(
            "formula-inline-prefixed",
            new CellSelector.ByAddress("Budget", "A2"),
            new MutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("=SUM(B1:B2)"))));
    MutationStep fileBacked =
        new MutationStep(
            "formula-file-backed",
            new CellSelector.ByAddress("Budget", "A3"),
            new MutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.utf8File("formula.txt"))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(stableInline, prefixedInline, fileBacked)),
            bindings);

    assertSame(stableInline, resolved.steps().get(0));

    MutationAction.SetCell prefixedAction =
        assertInstanceOf(
            MutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(1)).action());
    assertEquals(
        "SUM(B1:B2)",
        ((TextSourceInput.Inline)
                assertInstanceOf(CellInput.Formula.class, prefixedAction.value()).source())
            .text());

    MutationAction.SetCell fileBackedAction =
        assertInstanceOf(
            MutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(2)).action());
    assertEquals(
        "SUM(C1:C2)",
        ((TextSourceInput.Inline)
                assertInstanceOf(CellInput.Formula.class, fileBackedAction.value()).source())
            .text());
  }

  private static DrawingAnchorInput.TwoCell twoCellAnchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(0, 0, 0, 0),
        new DrawingMarkerInput(2, 3, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static PrintLayoutInput printLayoutWithHeader(HeaderFooterTextInput header) {
    PrintLayoutInput defaults = PrintLayoutInput.defaults();
    return new PrintLayoutInput(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        header,
        defaults.footer(),
        defaults.setup());
  }

  private static PrintLayoutInput printLayoutWithFooter(HeaderFooterTextInput footer) {
    PrintLayoutInput defaults = PrintLayoutInput.defaults();
    return new PrintLayoutInput(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        defaults.header(),
        footer,
        defaults.setup());
  }

  private static PrintLayoutInput printLayoutWithHeaderAndFooter(
      HeaderFooterTextInput header, HeaderFooterTextInput footer) {
    PrintLayoutInput defaults = PrintLayoutInput.defaults();
    return new PrintLayoutInput(
        defaults.printArea(),
        defaults.orientation(),
        defaults.scaling(),
        defaults.repeatingRows(),
        defaults.repeatingColumns(),
        header,
        footer,
        defaults.setup());
  }

  private static boolean requiresStandardInputFor(MutationAction action) {
    Object target =
        switch (action) {
          case MutationAction.SetCell _ -> new CellSelector.ByAddress("Budget", "A1");
          case MutationAction.SetRange _ -> new RangeSelector.ByRange("Budget", "A1:B2");
          case MutationAction.SetComment _ -> new CellSelector.ByAddress("Budget", "A1");
          case MutationAction.SetDataValidation _ -> new RangeSelector.ByRange("Budget", "A1");
          case MutationAction.SetTable setTable ->
              new dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet(
                  setTable.table().name(), setTable.table().sheetName());
          default -> new SheetSelector.ByName("Budget");
        };
    WorkbookPlan plan =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            mutations(mutate((dev.erst.gridgrind.contract.selector.Selector) target, action)));
    return SourceBackedPlanResolver.requiresStandardInput(plan);
  }

  private static PictureDataInput inlinePictureData() {
    return new PictureDataInput(
        ExcelPictureFormat.PNG,
        new BinarySourceInput.InlineBase64(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="));
  }

  private static ChartInput.Series chartSeries() {
    return new ChartInput.Series(
        new ChartInput.Title.None(),
        new ChartInput.DataSource.Reference("Budget!$A$2:$A$3"),
        new ChartInput.DataSource.Reference("Budget!$B$2:$B$3"),
        null,
        null,
        null,
        null);
  }

  private static ChartInput chartInput(
      String name,
      DrawingAnchorInput.TwoCell anchor,
      ChartInput.Title title,
      ChartInput.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      ChartInput.Plot plot) {
    return new ChartInput(
        name,
        anchor,
        title == null ? new ChartInput.Title.None() : title,
        legend == null
            ? new ChartInput.Legend.Visible(
                dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition.RIGHT)
            : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        List.of(plot));
  }

  private static byte[] pngBytes() {
    return Base64.getDecoder()
        .decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");
  }
}
