package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
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
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Source-backed canonical-structure preservation coverage. */
class SourceBackedPlanResolverCanonicalStructureCoverageTest
    extends SourceBackedPlanResolverTestSupport {
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
                    new CellMutationAction.SetCell(
                        new CellInput.RichText(
                            List.of(
                                new RichTextRunInput(
                                    new TextSourceInput.Utf8File("space.txt"), null)))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(plan, new ExecutionInputBindings(workingDirectory));

    CellMutationAction.SetCell richTextAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class,
            ((MutationStep) resolved.steps().getFirst()).action());
    CellInput.RichText richTextValue =
        assertInstanceOf(CellInput.RichText.class, richTextAction.value());
    assertEquals(" ", ((TextSourceInput.Inline) richTextValue.runs().getFirst().source()).text());
  }

  @Test
  void resolvePreservesCanonicalStepInstancesWhenNothingChanges() throws IOException {
    MutationStep mutationStep =
        new MutationStep(
            "mutate-inline",
            new CellSelector.ByAddress("Budget", "A1"),
            new CellMutationAction.SetCell(new CellInput.Text(TextSourceInput.inline("Owner"))));
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
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
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
            new CellMutationAction.SetRange(
                List.of(
                    List.of(
                        new CellInput.Text(TextSourceInput.inline("Owner")),
                        new CellInput.Numeric(42.0d)))));
    MutationStep commentStep =
        new MutationStep(
            "comment-inline",
            new CellSelector.ByAddress("Budget", "A1"),
            new CellMutationAction.SetComment(
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
            new DrawingMutationAction.SetPicture(
                new PictureInput(
                    "Logo", inlinePictureData(), twoCellAnchor(), TextSourceInput.inline("Logo"))));
    MutationStep shapeStep =
        new MutationStep(
            "shape-inline",
            new SheetSelector.ByName("Budget"),
            new DrawingMutationAction.SetShape(
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
            new StructuredMutationAction.SetDataValidation(
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
            new StructuredMutationAction.SetTable(
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
            new WorkbookMutationAction.SetPrintLayout(
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
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
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
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("=SUM(B2:B3)"))));
    MutationStep commentStep =
        new MutationStep(
            "comment-mixed",
            new CellSelector.ByAddress("Budget", "A1"),
            new CellMutationAction.SetComment(
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
            new DrawingMutationAction.SetPicture(
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
            new DrawingMutationAction.SetEmbeddedObject(
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
            new StructuredMutationAction.SetDataValidation(
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
            new StructuredMutationAction.SetDataValidation(
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
            new WorkbookMutationAction.SetPrintLayout(
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
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of(
                    formulaStep,
                    commentStep,
                    pictureStep,
                    embeddedObjectStep,
                    promptValidationStep,
                    errorValidationStep,
                    printLayoutStep)),
            new ExecutionInputBindings(workingDirectory));

    CellMutationAction.SetCell resolvedFormulaAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(0)).action());
    CellInput.Formula resolvedFormula =
        assertInstanceOf(CellInput.Formula.class, resolvedFormulaAction.value());
    assertEquals("SUM(B2:B3)", ((TextSourceInput.Inline) resolvedFormula.source()).text());

    CellMutationAction.SetComment resolvedCommentAction =
        assertInstanceOf(
            CellMutationAction.SetComment.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(1)).action());
    assertEquals(
        " Lovelace",
        ((TextSourceInput.Inline)
                resolvedCommentAction.comment().runs().orElseThrow().get(1).source())
            .text());

    DrawingMutationAction.SetPicture resolvedPictureAction =
        assertInstanceOf(
            DrawingMutationAction.SetPicture.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(2)).action());
    assertEquals(
        "cGF5bG9hZA==",
        ((BinarySourceInput.InlineBase64) resolvedPictureAction.picture().image().source())
            .base64Data());
    assertEquals(
        "Queue preview",
        ((TextSourceInput.Inline) resolvedPictureAction.picture().description()).text());

    DrawingMutationAction.SetEmbeddedObject resolvedEmbeddedAction =
        assertInstanceOf(
            DrawingMutationAction.SetEmbeddedObject.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(3)).action());
    assertEquals(
        Base64.getEncoder().encodeToString(pngBytes()),
        ((BinarySourceInput.InlineBase64)
                resolvedEmbeddedAction.embeddedObject().previewImage().source())
            .base64Data());

    StructuredMutationAction.SetDataValidation resolvedPromptValidation =
        assertInstanceOf(
            StructuredMutationAction.SetDataValidation.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(4)).action());
    assertEquals(
        "Pick one value",
        ((TextSourceInput.Inline) resolvedPromptValidation.validation().prompt().text()).text());

    StructuredMutationAction.SetDataValidation resolvedErrorValidation =
        assertInstanceOf(
            StructuredMutationAction.SetDataValidation.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(5)).action());
    assertEquals(
        "Try again",
        ((TextSourceInput.Inline) resolvedErrorValidation.validation().errorAlert().text()).text());

    WorkbookMutationAction.SetPrintLayout resolvedPrintLayout =
        assertInstanceOf(
            WorkbookMutationAction.SetPrintLayout.class,
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
            new CellMutationAction.SetCell(
                new CellInput.Text(TextSourceInput.utf8File("value.txt"))));
    MutationStep pictureDescriptionOnlyChanged =
        new MutationStep(
            "picture-description-only",
            new SheetSelector.ByName("Budget"),
            new DrawingMutationAction.SetPicture(
                new PictureInput(
                    "Logo",
                    inlinePictureData(),
                    twoCellAnchor(),
                    TextSourceInput.utf8File("value.txt"))));
    MutationStep footerOnlyChanged =
        new MutationStep(
            "footer-only-changed",
            new SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.SetPrintLayout(
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
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)"))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
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
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)"))));
    MutationStep prefixedInline =
        new MutationStep(
            "formula-inline-prefixed",
            new CellSelector.ByAddress("Budget", "A2"),
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("=SUM(B1:B2)"))));
    MutationStep fileBacked =
        new MutationStep(
            "formula-file-backed",
            new CellSelector.ByAddress("Budget", "A3"),
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.utf8File("formula.txt"))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of(stableInline, prefixedInline, fileBacked)),
            bindings);

    assertSame(stableInline, resolved.steps().get(0));

    CellMutationAction.SetCell prefixedAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(1)).action());
    assertEquals(
        "SUM(B1:B2)",
        ((TextSourceInput.Inline)
                assertInstanceOf(CellInput.Formula.class, prefixedAction.value()).source())
            .text());

    CellMutationAction.SetCell fileBackedAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(2)).action());
    assertEquals(
        "SUM(C1:C2)",
        ((TextSourceInput.Inline)
                assertInstanceOf(CellInput.Formula.class, fileBackedAction.value()).source())
            .text());
  }
}
