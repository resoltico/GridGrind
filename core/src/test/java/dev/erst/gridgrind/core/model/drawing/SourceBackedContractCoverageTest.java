package dev.erst.gridgrind.core.model.drawing;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Coverage for source-backed contract value objects and contexts. */
class SourceBackedContractCoverageTest {
  @Test
  void sourceFactoriesDefaultsAndResolveInputContextCoverFileAndStandardInputBranches() {
    assertEquals("inline", TextSourceSpec.inline("inline").text());
    assertEquals("note.txt", TextSourceSpec.utf8File("note.txt").path());
    assertInstanceOf(TextSourceSpec.Stdin.class, TextSourceSpec.standardInput());
    assertEquals(
        "path must not be blank",
        assertThrows(IllegalArgumentException.class, () -> TextSourceSpec.utf8File(" "))
            .getMessage());

    assertEquals("cGF5bG9hZA==", BinarySourceSpec.inlineBase64("cGF5bG9hZA==").base64Data());
    assertEquals("payload.bin", BinarySourceSpec.file("payload.bin").path());
    assertInstanceOf(BinarySourceSpec.Stdin.class, BinarySourceSpec.standardInput());
    assertEquals(
        "path must not be blank",
        assertThrows(IllegalArgumentException.class, () -> BinarySourceSpec.file(" "))
            .getMessage());
    assertEquals(
        "base64Data must be valid base64",
        assertThrows(
                IllegalArgumentException.class, () -> BinarySourceSpec.inlineBase64("not-base64"))
            .getMessage());

    ProblemContext.ResolveInputs context =
        new ProblemContext.ResolveInputs(
            ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE"),
            ProblemContextWorkbookSurfaces.InputReference.path("cell text", "missing.txt"));
    assertEquals("RESOLVE_INPUTS", context.type());
    assertEquals(java.util.Optional.of("cell text"), context.inputKind());
    assertEquals(java.util.Optional.of("missing.txt"), context.inputPath());
  }

  @Test
  void cellTextRejectsXml10ForbiddenCharactersWithoutReplacingThem() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new CellValueSpec.Text(TextSourceSpec.inline("A\u0007B")));

    assertEquals(
        "value.text contains XML 1.0-forbidden code point U+0007 at code-point index 1",
        failure.getMessage());
  }

  @Test
  void sourceBackedDtosPreserveDeferredValuesAndDefaultOptionalFlags() {
    CommentSpec defaultVisibleComment =
        CommentSpec.plain(TextSourceSpec.inline("Owner note"), "Ada", false);
    assertFalse(defaultVisibleComment.visible());

    CommentSpec fileBackedComment =
        new CommentSpec(
            TextSourceSpec.utf8File("comment.txt"),
            "Ada",
            false,
            java.util.Optional.of(
                java.util.List.of(
                    new RichTextRunSpec(
                        TextSourceSpec.utf8File("run.txt"), java.util.Optional.empty()))),
            java.util.Optional.empty());
    assertNotNull(fileBackedComment.runs());

    CommentSpec mixedSourceComment =
        new CommentSpec(
            TextSourceSpec.inline("Ada Lovelace"),
            "Ada",
            false,
            java.util.Optional.of(
                java.util.List.of(
                    new RichTextRunSpec(TextSourceSpec.inline("Ada"), java.util.Optional.empty()),
                    new RichTextRunSpec(
                        TextSourceSpec.utf8File("run-2.txt"), java.util.Optional.empty()))),
            java.util.Optional.empty());
    assertEquals(2, mixedSourceComment.runs().orElseThrow().size());

    HeaderFooterTextSpec defaults = HeaderFooterTextSpec.blank();
    assertEquals("", ((TextSourceSpec.Inline) defaults.left()).text());
    assertEquals("", ((TextSourceSpec.Inline) defaults.center()).text());
    assertEquals("", ((TextSourceSpec.Inline) defaults.right()).text());

    DataValidationPromptSpec prompt =
        new DataValidationPromptSpec(
            TextSourceSpec.utf8File("prompt-title.txt"), TextSourceSpec.standardInput(), true);
    assertTrue(prompt.showPromptBox());

    DataValidationErrorAlertSpec alert =
        new DataValidationErrorAlertSpec(
            ExcelDataValidationErrorStyle.STOP,
            TextSourceSpec.standardInput(),
            TextSourceSpec.utf8File("alert-text.txt"),
            true);
    assertTrue(alert.showErrorBox());

    ChartTitleSpec.Text chartTitle = new ChartTitleSpec.Text(TextSourceSpec.utf8File("title"));
    assertEquals("title", ((TextSourceSpec.Utf8File) chartTitle.source()).path());
  }

  @Test
  void sourceBackedMutationPayloadsRemainValidWithoutInlineCoercion() {
    DrawingAnchorSpec.TwoCell anchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(0, 0, 0, 0),
            new DrawingMarkerSpec(2, 2, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);

    PictureSpec picture =
        new PictureSpec(
            "Logo",
            new PictureDataSpec(ExcelPictureFormat.PNG, BinarySourceSpec.file("logo.png")),
            anchor,
            Optional.of(TextSourceSpec.utf8File("logo.txt")));
    assertEquals("logo.png", ((BinarySourceSpec.File) picture.image().source()).path());

    ShapeSpec.SimpleShape shape =
        new ShapeSpec.SimpleShape(
            "Banner", anchor, "rect", Optional.of(TextSourceSpec.standardInput()));
    assertInstanceOf(TextSourceSpec.Stdin.class, shape.text().orElseThrow());

    EmbeddedObjectSpec embeddedObject =
        new EmbeddedObjectSpec(
            "Payload",
            "Payload",
            "payload.bin",
            "open",
            BinarySourceSpec.standardInput(),
            new PictureDataSpec(ExcelPictureFormat.PNG, BinarySourceSpec.file("preview.png")),
            anchor);
    assertInstanceOf(BinarySourceSpec.Stdin.class, embeddedObject.payload());

    DataValidationSpec validation =
        new DataValidationSpec(
            new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.BETWEEN, "1", Optional.of("10")),
            false,
            false,
            Optional.of(
                new DataValidationPromptSpec(
                    TextSourceSpec.utf8File("prompt-title.txt"),
                    TextSourceSpec.utf8File("prompt.txt"),
                    true)),
            Optional.of(
                new DataValidationErrorAlertSpec(
                    ExcelDataValidationErrorStyle.STOP,
                    TextSourceSpec.utf8File("alert-title.txt"),
                    TextSourceSpec.utf8File("alert.txt"),
                    true)));
    assertTrue(validation.prompt().orElseThrow().showPromptBox());
    assertTrue(validation.errorAlert().orElseThrow().showErrorBox());
  }
}
