package dev.erst.gridgrind.core.model.style;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Validates canonical border patches independently of the workbook backend. */
class BorderPatchValidationTest {
  @Test
  void validatesBorderPatchRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new BorderSidePatch(null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    BorderSidePatch colorOnly = new BorderSidePatch(null, ColorSpec.rgb(new RgbHex("#a1b2c3")));
    assertEquals(Optional.empty(), colorOnly.style());
    assertEquals(Optional.of(ColorSpec.rgb(new RgbHex("#A1B2C3"))), colorOnly.color());
  }
}
