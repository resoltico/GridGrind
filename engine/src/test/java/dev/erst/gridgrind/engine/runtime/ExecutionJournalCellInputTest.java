package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import org.junit.jupiter.api.Test;

/** Journal presentation of authored raw-formula descriptors. */
class ExecutionJournalCellInputTest {
  @Test
  void rawFormulaLabelRetainsAuthoredText() {
    assertEquals(
        "RawFormula[text=LAMBDA(x,x+1)(A1)]",
        ExecutionJournalTargetResolver.summarizeCellInput(
            new CellValueSpec.RawFormula(TextSourceSpec.inline("LAMBDA(x,x+1)(A1)"))));
  }
}
