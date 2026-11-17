package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.List;
import org.junit.jupiter.api.Test;

/**
 * Independent container layouts distinguish complete streams, shared tables, aliases and cycles.
 */
class TiffContainerStructureTest {
  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void completeAndAbbreviatedJpegMembersSupportMultipleStripsAndTiles() throws Exception {
    for (String name :
        List.of(
            "jpeg7-strip-complete.tiff",
            "jpeg7-strip-shared-tables.tiff",
            "jpeg7-tile-complete.tiff",
            "jpeg7-tile-shared-tables.tiff")) {
      var geometry = read(name, new MediaAdmissionBudget());
      assertEquals(17, geometry.width());
      assertEquals(19, geometry.height());
    }
    var failure =
        assertThrows(
            InvalidRequestException.class,
            () -> read("jpeg7-malformed-tables.tiff", new MediaAdmissionBudget()));
    var field = (FieldValidationProblem) failure.requestProblem();
    assertEquals(FieldValidationMediaRule.FORMAT_STRUCTURE, field.rule());
    assertEquals("HEADER", field.operands().get(1));
  }

  @Test
  void repeatedChildAliasesAreDistinctFromAnActiveGraphCycle() throws Exception {
    var budget = new MediaAdmissionBudget();
    assertEquals(1, read("aliased-child.tiff", budget).width());
    // Two image directories consume two pixels; a repeated edge does not create another image.
    assertDoesNotThrow(() -> budget.raster(16_777_214, 1));
    assertThrows(InvalidRequestException.class, () -> budget.raster(1, 1));
    var failure =
        assertThrows(
            InvalidRequestException.class,
            () -> read("active-cycle.tiff", new MediaAdmissionBudget()));
    assertEquals("OFFSET", ((FieldValidationProblem) failure.requestProblem()).operands().get(1));
  }

  @Test
  void directoryDepthIncludesTheExactLimitAndRefusesTheNextChild() throws Exception {
    assertDoesNotThrow(() -> read("depth-8.tiff", new MediaAdmissionBudget()));
    var failure =
        assertThrows(
            InvalidRequestException.class, () -> read("depth-9.tiff", new MediaAdmissionBudget()));
    assertEquals(
        FieldValidationMediaRule.RESOURCE_LIMIT,
        ((FieldValidationProblem) failure.requestProblem()).rule());
    assertEquals(
        "LIM-048", ((FieldValidationProblem) failure.requestProblem()).operands().getFirst());
  }

  @Test
  void codingRowsDoNotInventASeparateImageOrContainerRecordCap() throws Exception {
    var geometry = read("tall-single-frame.tiff", new MediaAdmissionBudget());
    assertEquals(1, geometry.width());
    assertEquals(65_536, geometry.height());
  }

  private static ImageGeometry read(String name, MediaAdmissionBudget budget)
      throws java.io.IOException {
    byte[] bytes = AuthoredImageAdmissionTest.fixture(name);
    return ImageStructures.check(
        new MediaInput(bytes, 0, bytes.length, ExcelPictureFormat.TIFF, budget, 0));
  }
}
