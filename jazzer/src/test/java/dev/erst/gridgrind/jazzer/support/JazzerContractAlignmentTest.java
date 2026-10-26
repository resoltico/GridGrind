package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import org.junit.jupiter.api.Test;

/** Guards Jazzer support helpers against contract-model drift. */
class JazzerContractAlignmentTest {
  @Test
  void nextProtocolGridProducesCanonicalTypedGridInput() {
    CellGridSpec.Typed gridInput =
        assertInstanceOf(
            CellGridSpec.Typed.class,
            FuzzDataDecoders.nextProtocolGrid(GridGrindFuzzData.replay(new byte[0]), 2, 2));

    assertEquals(2, gridInput.cells().size());
    assertEquals(2, gridInput.cells().getFirst().size());
    assertEquals(new CellValueSpec.Blank(), gridInput.cells().getFirst().getFirst());
  }

  @Test
  void nextProtocolRowProducesCanonicalTypedRowInput() {
    CellRowSpec.Typed rowInput =
        assertInstanceOf(
            CellRowSpec.Typed.class,
            FuzzDataDecoders.nextProtocolRow(GridGrindFuzzData.replay(new byte[0]), 3));

    assertEquals(3, rowInput.cells().size());
    assertEquals(new CellValueSpec.Blank(), rowInput.cells().getFirst());
  }

  @Test
  void errorCellsRemainValueBearingForRoundTripFootprints() {
    assertTrue(
        XlsxRoundTripExpectedFootprintSupport.isValueBearing(new CellValue.ErrorValue("#REF!")));
  }
}
