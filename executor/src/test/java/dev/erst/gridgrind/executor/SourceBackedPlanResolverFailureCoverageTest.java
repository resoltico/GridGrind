package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Source-backed resolver failure and predicate coverage. */
class SourceBackedPlanResolverFailureCoverageTest extends SourceBackedPlanResolverTestSupport {
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
                    new CellMutationAction.SetCell(
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
                    new CellMutationAction.SetCell(
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
                    new CellMutationAction.SetCell(
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
                    new CellMutationAction.SetCell(
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
                    new CellMutationAction.SetCell(
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
                    new DrawingMutationAction.SetPicture(
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
                    new DrawingMutationAction.SetPicture(
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
                    new DrawingMutationAction.SetEmbeddedObject(
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

    assertFalse(SourceBackedInputRequirements.requiresStandardInput(new ChartTitleInput.None()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartTitleInput.Formula("Budget!$A$1")));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartTitleInput.Text(TextSourceInput.inline("Title"))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new ChartTitleInput.Text(TextSourceInput.standardInput())));

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
                    new CellMutationAction.SetCell(
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
                    new CellMutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.StandardInput())))));
    InputSourceUnavailableException unavailable =
        assertThrows(
            InputSourceUnavailableException.class,
            () ->
                SourceBackedPlanResolver.resolve(
                    standardInputPlan, new ExecutionInputBindings(Path.of(""))));
    assertEquals("cell text", unavailable.inputKind());
  }
}
