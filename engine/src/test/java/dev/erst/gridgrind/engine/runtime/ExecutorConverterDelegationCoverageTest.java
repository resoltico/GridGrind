package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.engine.backend.xssf.*;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers delegation-only converter wrappers introduced by executor seam refactors. */
class ExecutorConverterDelegationCoverageTest {
  @Test
  void workbookCommandConverterDelegatesRemainCovered() {
    RichTextRunSpec run =
        new RichTextRunSpec(TextSourceSpec.inline("Bold"), java.util.Optional.empty());

    assertEquals(
        "Bold",
        MutationResolution.toResolvedRichText(new CellValueSpec.RichText(List.of(run)))
            .runs()
            .getFirst()
            .text());
    assertEquals("Bold", MutationResolution.toResolvedRichTextRun(run).text());
  }
}
