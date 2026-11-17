package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Source-free preparation refusals, optional operands and target validity. */
class MutationPreparationBoundaryTest {
  @Test
  void converterNullBranchesAndResolvedPromptFieldsStayTyped() {
    assertTrue(MutationPreparation.toResolvedValidationPrompt(null).isEmpty());
    assertTrue(MutationPreparation.toResolvedValidationErrorAlert(null).isEmpty());

    Optional<ResolvedValidationPrompt> prompt =
        MutationPreparation.toResolvedValidationPrompt(
            new DataValidationPromptSpec(
                TextSourceSpec.inline("Prompt title"), TextSourceSpec.inline("Prompt text"), true));
    Optional<ResolvedValidationErrorAlert> alert =
        MutationPreparation.toResolvedValidationErrorAlert(
            new DataValidationErrorAlertSpec(
                ExcelDataValidationErrorStyle.WARNING,
                TextSourceSpec.inline("Alert title"),
                TextSourceSpec.inline("Alert text"),
                true));

    assertEquals("Prompt title", prompt.orElseThrow().title());
    assertEquals(ExcelDataValidationErrorStyle.WARNING, alert.orElseThrow().style());
  }

  @Test
  void rejectsUnresolvedTextAndPictureOperands() {
    assertEquals(
        "cell text must be resolved to INLINE before conversion",
        assertThrows(
                IllegalStateException.class,
                () ->
                    MutationPreparation.toResolvedCellWrite(
                        new CellValueSpec.Text(TextSourceSpec.utf8File("title.txt"))))
            .getMessage());

    PictureSpec unresolvedPicture =
        new PictureSpec(
            "Logo",
            new PictureDataSpec(ExcelPictureFormat.PNG, BinarySourceSpec.file("logo.png")),
            new dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec.TwoCell(
                new dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec(0, 0, 0, 0),
                new dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec(1, 1, 0, 0),
                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
            Optional.of(TextSourceSpec.inline("Logo")));
    assertEquals(
        "picture payload must be resolved to INLINE_BASE64 before conversion",
        assertThrows(
                IllegalStateException.class,
                () -> MutationPreparation.toResolvedPicture(unresolvedPicture))
            .getMessage());
  }

  @Test
  void conditionalFormattingMutationRequiresExplicitRangeTargets() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ResolvedStep(
                    new RangeSelector.AllOnSheet(new SheetName("Budget")),
                    new StructuredMutationAction.SetConditionalFormatting(
                        new ConditionalFormattingSpec(
                            List.of(
                                new ConditionalFormattingRuleSpec.FormulaRule(
                                    "A1>0", true, Optional.empty()))))));

    assertEquals("Mutation target does not match SetConditionalFormatting", failure.getMessage());
  }

  @Test
  void storedErrorPreparationRetainsItsCanonicalValue() {
    CellValueSpec.ErrorValue errorValue = new CellValueSpec.ErrorValue("#REF!");
    assertEquals(
        "#REF!",
        org.junit.jupiter.api.Assertions.assertInstanceOf(
                CellValueSpec.ErrorValue.class, MutationPreparation.toResolvedCellWrite(errorValue))
            .value());
  }
}
