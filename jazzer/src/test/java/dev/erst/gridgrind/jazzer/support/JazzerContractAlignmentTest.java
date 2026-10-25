package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.value.CellGridInput;
import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.model.value.CellRowInput;
import dev.erst.gridgrind.excel.ExcelCellValue;
import org.junit.jupiter.api.Test;

/** Guards Jazzer support helpers against contract-model drift. */
class JazzerContractAlignmentTest {
  @Test
  void nextProtocolGridProducesCanonicalTypedGridInput() {
    CellGridInput.Typed gridInput =
        assertInstanceOf(
            CellGridInput.Typed.class,
            FuzzDataDecoders.nextProtocolGrid(GridGrindFuzzData.replay(new byte[0]), 2, 2));

    assertEquals(2, gridInput.cells().size());
    assertEquals(2, gridInput.cells().getFirst().size());
    assertEquals(new CellInput.Blank(), gridInput.cells().getFirst().getFirst());
  }

  @Test
  void nextProtocolRowProducesCanonicalTypedRowInput() {
    CellRowInput.Typed rowInput =
        assertInstanceOf(
            CellRowInput.Typed.class,
            FuzzDataDecoders.nextProtocolRow(GridGrindFuzzData.replay(new byte[0]), 3));

    assertEquals(3, rowInput.cells().size());
    assertEquals(new CellInput.Blank(), rowInput.cells().getFirst());
  }

  @Test
  void errorCellsRemainValueBearingForRoundTripFootprints() {
    assertTrue(XlsxRoundTripExpectedFootprintSupport.isValueBearing(ExcelCellValue.error("#REF!")));
  }
}
