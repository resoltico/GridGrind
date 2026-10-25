package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.drawing.CommentInput;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectInput;
import dev.erst.gridgrind.core.model.drawing.PictureDataInput;
import dev.erst.gridgrind.core.model.formula.InvalidRawFormulaTextException;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextInput;
import dev.erst.gridgrind.core.model.style.RichTextRunInput;
import dev.erst.gridgrind.core.model.table.TableInput;
import dev.erst.gridgrind.core.model.table.TableStyleInput;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertInput;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptInput;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleInput;
import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceInput;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(
                        new CellInput.RichText(
                            List.of(
                                new RichTextRunInput(
                                    new TextSourceInput.Utf8File("space.txt"),
                                    java.util.Optional.empty())))))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory, List.of(workingDirectory.resolve("space.txt"))));

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
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-inline-steps-");
    MutationStep mutationStep =
        new MutationStep(
            new StepId("mutate-inline"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.SetCell(new CellInput.Text(TextSourceInput.inline("Owner"))));
    InspectionStep inspectionStep =
        new InspectionStep(
            new StepId("inspect-inline"),
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByName("BudgetTable"),
                    "Item",
                    new CellInput.Text(TextSourceInput.inline("Hosting"))),
                "Amount"),
            new SheetIntrospectionQuery.GetCells());
    AssertionStep assertionStep =
        new AssertionStep(
            new StepId("assert-inline"),
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByName("BudgetTable"),
                    "Item",
                    new CellInput.Text(TextSourceInput.inline("Hosting"))),
                "Amount"),
            new CellAssertion.CellValue(
                new dev.erst.gridgrind.core.model.value.CellScalarValue.NumberValue(125.0)));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.core.plan.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput.empty(),
                List.of(mutationStep, inspectionStep, assertionStep)),
            ExecutionInputBindingsFixtureSupport.bindings(workingDirectory));

    assertSame(mutationStep, resolved.steps().get(0));
    assertSame(inspectionStep, resolved.steps().get(1));
    assertSame(assertionStep, resolved.steps().get(2));
  }

  @Test
  void resolvePreservesAlreadyInlineMutationFamiliesWhenNothingNeedsInlining() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-source-backed-inline-families-");
    MutationStep rangeStep =
        new MutationStep(
            new StepId("range-inline"),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.core.model.value.CellGridInput.Typed(
                    List.of(
                        List.of(
                            new CellInput.Text(TextSourceInput.inline("Owner")),
                            new CellInput.NumberValue(42.0d))))));
    MutationStep commentStep =
        new MutationStep(
            new StepId("comment-inline"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada"),
                    "Ada",
                    false,
                    java.util.Optional.empty(),
                    java.util.Optional.empty())));
    MutationStep pictureStep =
        new MutationStep(
            new StepId("picture-inline"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new DrawingMutationAction.SetPicture(
                picture(
                    "Logo", inlinePictureData(), twoCellAnchor(), TextSourceInput.inline("Logo"))));
    MutationStep shapeStep =
        new MutationStep(
            new StepId("shape-inline"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new DrawingMutationAction.SetShape(
                simpleShape(
                    "Shape", twoCellAnchor(), "roundRect", TextSourceInput.inline("Shape"))));
    MutationStep validationStep =
        new MutationStep(
            new StepId("validation-inline"),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1")),
            new StructuredMutationAction.SetDataValidation(
                dataValidation(
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
            new StepId("table-inline"),
            new TableSelector.ByNameOnSheet("BudgetTable", new SheetName("Budget")),
            new StructuredMutationAction.SetTable(
                new TableInput(
                    "BudgetTable",
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new RangeRef("A1:B2"),
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
            new StepId("print-layout-inline"),
            new SheetSelector.ByName(new SheetName("Budget")),
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
                dev.erst.gridgrind.core.plan.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput.empty(),
                List.of(
                    rangeStep,
                    commentStep,
                    pictureStep,
                    shapeStep,
                    validationStep,
                    tableStep,
                    printLayoutStep)),
            ExecutionInputBindingsFixtureSupport.bindings(workingDirectory));

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
            new StepId("formula-inline-prefix"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(B2:B3)"))));
    MutationStep commentStep =
        new MutationStep(
            new StepId("comment-mixed"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.SetComment(
                new CommentInput(
                    TextSourceInput.inline("Ada Lovelace"),
                    "Ada",
                    false,
                    java.util.Optional.of(
                        List.of(
                            new RichTextRunInput(
                                TextSourceInput.inline("Ada"), java.util.Optional.empty()),
                            richTextRun(TextSourceInput.utf8File("comment-run.txt"), null))),
                    java.util.Optional.empty())));
    MutationStep pictureStep =
        new MutationStep(
            new StepId("picture-mixed"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new DrawingMutationAction.SetPicture(
                picture(
                    "Logo",
                    new PictureDataInput(
                        ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64("cGF5bG9hZA")),
                    twoCellAnchor(),
                    TextSourceInput.utf8File("picture-description.txt"))));
    MutationStep embeddedObjectStep =
        new MutationStep(
            new StepId("embedded-mixed"),
            new SheetSelector.ByName(new SheetName("Budget")),
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
            new StepId("validation-prompt-mixed"),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1")),
            new StructuredMutationAction.SetDataValidation(
                dataValidation(
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
            new StepId("validation-error-mixed"),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("B1")),
            new StructuredMutationAction.SetDataValidation(
                dataValidation(
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
            new StepId("print-layout-mixed"),
            new SheetSelector.ByName(new SheetName("Budget")),
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

    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput.empty(),
            List.of(
                formulaStep,
                commentStep,
                pictureStep,
                embeddedObjectStep,
                promptValidationStep,
                errorValidationStep,
                printLayoutStep));
    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory,
                List.of(
                    workingDirectory.resolve("comment-run.txt"),
                    workingDirectory.resolve("picture-description.txt"),
                    workingDirectory.resolve("preview.bin"),
                    workingDirectory.resolve("prompt-body.txt"),
                    workingDirectory.resolve("error-body.txt"),
                    workingDirectory.resolve("header-center.txt"))));

    CellMutationAction.SetCell resolvedFormulaAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(0)).action());
    CellInput.Formula resolvedFormula =
        assertInstanceOf(CellInput.Formula.class, resolvedFormulaAction.value());
    assertEquals(
        "SUM(B2:B3)",
        assertInstanceOf(TextSourceInput.Inline.class, resolvedFormula.source()).text());

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
        ((TextSourceInput.Inline) resolvedPictureAction.picture().description().orElseThrow())
            .text());

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
        ((TextSourceInput.Inline)
                resolvedPromptValidation.validation().prompt().orElseThrow().text())
            .text());

    StructuredMutationAction.SetDataValidation resolvedErrorValidation =
        assertInstanceOf(
            StructuredMutationAction.SetDataValidation.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(5)).action());
    assertEquals(
        "Try again",
        ((TextSourceInput.Inline)
                resolvedErrorValidation.validation().errorAlert().orElseThrow().text())
            .text());

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
            new StepId("target-and-action-changed"),
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
            new StepId("picture-description-only"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new DrawingMutationAction.SetPicture(
                picture(
                    "Logo",
                    inlinePictureData(),
                    twoCellAnchor(),
                    TextSourceInput.utf8File("value.txt"))));
    MutationStep footerOnlyChanged =
        new MutationStep(
            new StepId("footer-only-changed"),
            new SheetSelector.ByName(new SheetName("Budget")),
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
            new StepId("formula-inline-stable"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B2")),
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)"))));

    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput.empty(),
            List.of(
                targetAndActionChanged,
                pictureDescriptionOnlyChanged,
                footerOnlyChanged,
                formulaAlreadyInline));
    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory,
                List.of(
                    workingDirectory.resolve("item.txt"),
                    workingDirectory.resolve("value.txt"),
                    workingDirectory.resolve("footer-right.txt"))));

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
        workingDirectory.resolve("formula.txt"), "SUM(C1:C2)", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("raw-formula.txt"), "LAMBDA(x,x+1)(A1)", StandardCharsets.UTF_8);
    Files.writeString(
        workingDirectory.resolve("invalid-raw-formula.txt"),
        "=LAMBDA(x,x+1)(A1)",
        StandardCharsets.UTF_8);
    ExecutionInputBindings bindings =
        ExecutionInputBindingsFixtureSupport.bindings(
            workingDirectory,
            List.of(
                workingDirectory.resolve("formula.txt"),
                workingDirectory.resolve("raw-formula.txt"),
                workingDirectory.resolve("invalid-raw-formula.txt")));

    TextSourceInput.Inline stableSource = TextSourceInput.inline("SUM(A1:A2)");
    assertSame(stableSource, SourceBackedPlanResolver.resolveFormulaSource(stableSource, bindings));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            SourceBackedPlanResolver.resolveFormulaSource(
                TextSourceInput.inline("=SUM(B1:B2)"), bindings));

    assertEquals(
        "SUM(C1:C2)",
        ((TextSourceInput.Inline)
                SourceBackedPlanResolver.resolveFormulaSource(
                    TextSourceInput.utf8File("formula.txt"), bindings))
            .text());
    assertEquals(
        "LAMBDA(x,x+1)(A1)",
        ((TextSourceInput.Inline)
                SourceBackedPlanResolver.resolveRawFormulaSource(
                    TextSourceInput.utf8File("raw-formula.txt"), bindings))
            .text());
    assertThrows(
        InvalidRawFormulaTextException.class,
        () ->
            SourceBackedPlanResolver.resolveRawFormulaSource(
                TextSourceInput.utf8File("invalid-raw-formula.txt"), bindings));
    CellInput.RawFormula inlineRawFormula =
        new CellInput.RawFormula(TextSourceInput.inline("LAMBDA(x,x+1)(A1)"));
    assertSame(
        inlineRawFormula, SourceBackedPlanResolver.resolveCellInput(inlineRawFormula, bindings));
    assertEquals(
        "LAMBDA(x,x+1)(A1)",
        assertInstanceOf(
                TextSourceInput.Inline.class,
                assertInstanceOf(
                        CellInput.RawFormula.class,
                        SourceBackedPlanResolver.resolveCellInput(
                            new CellInput.RawFormula(TextSourceInput.utf8File("raw-formula.txt")),
                            bindings))
                    .source())
            .text());

    MutationStep stableInline =
        new MutationStep(
            new StepId("formula-inline-stable"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(A1:A2)"))));
    MutationStep secondInline =
        new MutationStep(
            new StepId("formula-inline-prefixed"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A2")),
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.inline("SUM(B1:B2)"))));
    MutationStep fileBacked =
        new MutationStep(
            new StepId("formula-file-backed"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A3")),
            new CellMutationAction.SetCell(
                new CellInput.Formula(TextSourceInput.utf8File("formula.txt"))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.core.plan.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput.empty(),
                List.of(stableInline, secondInline, fileBacked)),
            bindings);

    assertSame(stableInline, resolved.steps().get(0));

    CellMutationAction.SetCell prefixedAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(1)).action());
    assertEquals(
        "SUM(B1:B2)",
        assertInstanceOf(
                TextSourceInput.Inline.class,
                assertInstanceOf(CellInput.Formula.class, prefixedAction.value()).source())
            .text());

    CellMutationAction.SetCell fileBackedAction =
        assertInstanceOf(
            CellMutationAction.SetCell.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(2)).action());
    assertEquals(
        "SUM(C1:C2)",
        assertInstanceOf(
                TextSourceInput.Inline.class,
                assertInstanceOf(CellInput.Formula.class, fileBackedAction.value()).source())
            .text());
  }
}
