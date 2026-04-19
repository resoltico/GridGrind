package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import org.junit.jupiter.api.Test;

/** Dedicated coverage for Phase 7 source-backed contract value objects and contexts. */
class Phase7ContractCoverageTest {
  @Test
  void sourceFactoriesDefaultsAndResolveInputContextCoverFileAndStandardInputBranches() {
    assertEquals("inline", TextSourceInput.inline("inline").text());
    assertEquals("note.txt", TextSourceInput.utf8File("note.txt").path());
    assertInstanceOf(TextSourceInput.StandardInput.class, TextSourceInput.standardInput());
    assertEquals(
        "path must not be blank",
        assertThrows(IllegalArgumentException.class, () -> TextSourceInput.utf8File(" "))
            .getMessage());

    assertEquals("cGF5bG9hZA==", BinarySourceInput.inlineBase64("cGF5bG9hZA==").base64Data());
    assertEquals("payload.bin", BinarySourceInput.file("payload.bin").path());
    assertInstanceOf(BinarySourceInput.StandardInput.class, BinarySourceInput.standardInput());
    assertEquals(
        "path must not be blank",
        assertThrows(IllegalArgumentException.class, () -> BinarySourceInput.file(" "))
            .getMessage());
    assertEquals(
        "base64Data must be valid base64",
        assertThrows(
                IllegalArgumentException.class, () -> BinarySourceInput.inlineBase64("not-base64"))
            .getMessage());

    GridGrindResponse.ProblemContext.ResolveInputs context =
        new GridGrindResponse.ProblemContext.ResolveInputs("NEW", "NONE", "cell text", null);
    assertEquals("RESOLVE_INPUTS", context.stage());
    assertEquals("cell text", context.inputKind());
  }

  @Test
  void sourceBackedDtosPreserveDeferredValuesAndDefaultOptionalFlags() {
    CommentInput fileBackedComment =
        new CommentInput(
            TextSourceInput.utf8File("comment.txt"),
            "Ada",
            null,
            java.util.List.of(new RichTextRunInput(TextSourceInput.utf8File("run.txt"), null)),
            null);
    assertNotNull(fileBackedComment.runs());

    CommentInput mixedSourceComment =
        new CommentInput(
            TextSourceInput.inline("Ada Lovelace"),
            "Ada",
            false,
            java.util.List.of(
                new RichTextRunInput(TextSourceInput.inline("Ada"), null),
                new RichTextRunInput(TextSourceInput.utf8File("run-2.txt"), null)),
            null);
    assertEquals(2, mixedSourceComment.runs().size());

    HeaderFooterTextInput defaults = new HeaderFooterTextInput(null, null, null);
    assertEquals("", ((TextSourceInput.Inline) defaults.left()).text());
    assertEquals("", ((TextSourceInput.Inline) defaults.center()).text());
    assertEquals("", ((TextSourceInput.Inline) defaults.right()).text());

    DataValidationPromptInput prompt =
        new DataValidationPromptInput(
            TextSourceInput.utf8File("prompt-title.txt"), TextSourceInput.standardInput(), null);
    assertTrue(prompt.showPromptBox());

    DataValidationErrorAlertInput alert =
        new DataValidationErrorAlertInput(
            ExcelDataValidationErrorStyle.STOP,
            TextSourceInput.standardInput(),
            TextSourceInput.utf8File("alert-text.txt"),
            null);
    assertTrue(alert.showErrorBox());

    ChartInput.Title.Text chartTitle = new ChartInput.Title.Text(TextSourceInput.utf8File("title"));
    assertEquals("title", ((TextSourceInput.Utf8File) chartTitle.source()).path());
  }

  @Test
  void sourceBackedMutationPayloadsRemainValidWithoutInlineCoercion() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(0, 0, 0, 0),
            new DrawingMarkerInput(2, 2, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);

    PictureInput picture =
        new PictureInput(
            "Logo",
            new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.file("logo.png")),
            anchor,
            TextSourceInput.utf8File("logo.txt"));
    assertEquals("logo.png", ((BinarySourceInput.File) picture.image().source()).path());

    ShapeInput shape =
        new ShapeInput(
            "Banner",
            ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
            anchor,
            null,
            TextSourceInput.standardInput());
    assertInstanceOf(TextSourceInput.StandardInput.class, shape.text());

    EmbeddedObjectInput embeddedObject =
        new EmbeddedObjectInput(
            "Payload",
            "Payload",
            "payload.bin",
            "open",
            BinarySourceInput.standardInput(),
            new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.file("preview.png")),
            anchor);
    assertInstanceOf(BinarySourceInput.StandardInput.class, embeddedObject.payload());

    DataValidationInput validation =
        new DataValidationInput(
            new DataValidationRuleInput.WholeNumber(ExcelComparisonOperator.BETWEEN, "1", "10"),
            null,
            null,
            new DataValidationPromptInput(
                TextSourceInput.utf8File("prompt-title.txt"),
                TextSourceInput.utf8File("prompt.txt"),
                null),
            new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceInput.utf8File("alert-title.txt"),
                TextSourceInput.utf8File("alert.txt"),
                null));
    assertTrue(validation.prompt().showPromptBox());
    assertTrue(validation.errorAlert().showErrorBox());
  }
}
