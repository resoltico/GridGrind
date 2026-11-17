package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Prepares rich text and individual rich-text runs without engine dependencies. */
class RichTextPreparationTest {
  @Test
  void richTextAndIndividualRunsPreserveText() {
    RichTextRunSpec run =
        new RichTextRunSpec(TextSourceSpec.inline("Bold"), java.util.Optional.empty());

    assertEquals(
        "Bold",
        MutationPreparation.toResolvedRichText(new CellValueSpec.RichText(List.of(run)))
            .runs()
            .getFirst()
            .text());
    assertEquals("Bold", MutationPreparation.toResolvedRichTextRun(run).text());
  }
}
