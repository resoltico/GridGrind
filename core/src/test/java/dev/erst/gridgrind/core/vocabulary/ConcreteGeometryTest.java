package dev.erst.gridgrind.core.vocabulary;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Checks canonical coordinate identity, worksheet bounds and ordered rectangle geometry. */
class ConcreteGeometryTest {
  @ParameterizedTest
  @ValueSource(strings = {"A1", "a1", "$A$1", "A$1", "$A1"})
  void positionsIgnoreCaseAndAnchors(String alias) {
    var address = new CellAddress(alias);
    assertEquals(new CellAddress(0, 0), address);
    assertEquals("A1", address.value());
  }

  @Test
  void positionsRetainIndependentIndices() {
    assertEquals("AA10", new CellAddress(9, 26).value());
    assertEquals(new CellAddress(9, 26), new CellAddress("$AA$10"));
    assertEquals(new CellAddress(1048575, 16383), new CellAddress("$xfd$1048576"));
  }

  @ParameterizedTest
  @ValueSource(strings = {"XFE1", "A1048577", "A0", "A01", "1", "A1:B2", "A99999999999999999"})
  void rejectsMalformedOrOutOfBoundsPositions(String invalid) {
    assertThrows(IllegalArgumentException.class, () -> new CellAddress(invalid));
  }

  @Test
  void rectanglesHaveOneIdentityAndExactGeometry() {
    assertEquals(new RangeRef("A1"), new RangeRef("$a$1:A1"));
    RangeRef rectangle = new RangeRef("$B3:D$5");
    assertEquals(new RangeRef(2, 4, 1, 3), rectangle);
    assertEquals(3, rectangle.rowCount());
    assertEquals(3, rectangle.columnCount());
    assertEquals(9, rectangle.cellCount());
    assertTrue(rectangle.contains(new CellAddress("C4")));
    assertFalse(rectangle.contains(new CellAddress("A4")));
    assertTrue(rectangle.intersects(new RangeRef("D5:F7")));
    assertFalse(rectangle.intersects(new RangeRef("E5:F7")));
    assertEquals(17179869184L, new RangeRef("A1:XFD1048576").cellCount());
    var invalidEndpoint =
        assertThrows(
            dev.erst.gridgrind.core.problem.InvalidRequestException.class,
            () -> new RangeRef("A1:"));
    var originalAddress =
        assertInstanceOf(
            dev.erst.gridgrind.core.problem.InvalidRequestException.class,
            invalidEndpoint.getCause());
    assertEquals(
        "range",
        assertInstanceOf(
                dev.erst.gridgrind.core.problem.FieldValidationProblem.class,
                invalidEndpoint.requestProblem())
            .fieldName());
    assertEquals(
        "address",
        assertInstanceOf(
                dev.erst.gridgrind.core.problem.FieldValidationProblem.class,
                originalAddress.requestProblem())
            .fieldName());
    assertThrows(IllegalArgumentException.class, () -> new RangeRef(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new RangeRef(0, 1048576, 0, 0));
  }

  @ParameterizedTest
  @ValueSource(strings = {"B2:A1", "A2:B1", "B1:A2", "A1:", "A1:B2:C3", "A1:?", " "})
  void rejectsMalformedOrReversedRectangles(String invalid) {
    assertThrows(IllegalArgumentException.class, () -> new RangeRef(invalid));
  }

  @Test
  void selectorListsRejectAliasesButRetainOrderedOverlaps() {
    SheetName sheet = new SheetName("Sheet");
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellSelector.ByAddresses(
                sheet, List.of(new CellAddress("A1"), new CellAddress("$a$1"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RangeSelector.ByRanges(sheet, List.of(new RangeRef("A1"), new RangeRef("A1:A1"))));
    var overlapping = List.of(new RangeRef("A1:B2"), new RangeRef("B2:C3"));
    assertEquals(overlapping, new RangeSelector.ByRanges(sheet, overlapping).ranges());
    assertEquals("$B$2:A1", new RangeLabel("$B$2:A1").value());
  }
}
