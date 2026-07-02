package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.math.BigDecimal;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Coverage tests for schema type inference branches in {@link ExcelWorkbookIntrospector}. */
class ExcelWorkbookIntrospectorCoverageTest {
  @Test
  void schemaObservedTypeReflectsTemporalProjectionAndRejectsFormulaSnapshots() {
    ExcelCellReadProjection temporalProjection =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.TEMPORAL));
    ExcelCellReadProjection defaultProjection = ExcelCellReadProjection.defaults();
    ExcelCellSnapshot.NumberSnapshot dateCell = numberSnapshot("A1", 45839.0d, "m/d/yyyy");
    ExcelCellSnapshot.NumberSnapshot dateTimeCell = numberSnapshot("A2", 45839.5d, "m/d/yyyy h:mm");
    ExcelCellSnapshot.NumberSnapshot timeCell = numberSnapshot("A2b", 0.5d, "h:mm");
    ExcelCellSnapshot.NumberSnapshot plainNumber = numberSnapshot("A3", 42.0d, "0.00");
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot("A4", "", style("General"), metadata());
    ExcelCellSnapshot.FormulaSnapshot formula =
        new ExcelCellSnapshot.FormulaSnapshot(
            "A5", "42", style("General"), metadata(), "SUM(A1:A3)", plainNumber);

    assertEquals("DATE", schemaObservedType(dateCell, temporalProjection));
    assertEquals("TIME", schemaObservedType(timeCell, temporalProjection));
    assertEquals("DATE_TIME", schemaObservedType(dateTimeCell, temporalProjection));
    assertEquals("NUMBER", schemaObservedType(plainNumber, temporalProjection));
    assertEquals("NUMBER", schemaObservedType(plainNumber, defaultProjection));
    assertEquals("BLANK", schemaObservedType(blank, defaultProjection));

    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class, () -> schemaObservedType(formula, defaultProjection));
    assertEquals("Schema observed types must not receive FORMULA directly", exception.getMessage());
  }

  private static String schemaObservedType(
      ExcelCellSnapshot snapshot, ExcelCellReadProjection projection) {
    try {
      return (String) schemaObservedTypeMethod().invoke(snapshot, projection);
    } catch (RuntimeException runtimeException) {
      throw runtimeException;
    } catch (Throwable throwable) {
      throw new AssertionError(throwable);
    }
  }

  private static MethodHandle schemaObservedTypeMethod() {
    try {
      return MethodHandles.privateLookupIn(ExcelWorkbookIntrospector.class, MethodHandles.lookup())
          .findStatic(
              ExcelWorkbookIntrospector.class,
              "schemaObservedType",
              MethodType.methodType(
                  String.class, ExcelCellSnapshot.class, ExcelCellReadProjection.class));
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError(exception.getMessage(), exception);
    }
  }

  private static ExcelCellSnapshot.NumberSnapshot numberSnapshot(
      String address, double numberValue, String numberFormat) {
    return new ExcelCellSnapshot.NumberSnapshot(
        address, Double.toString(numberValue), style(numberFormat), metadata(), numberValue);
  }

  private static ExcelCellMetadataSnapshot metadata() {
    return ExcelCellMetadataSnapshot.empty();
  }

  private static ExcelCellStyleSnapshot style(String numberFormat) {
    ExcelBorderSideSnapshot emptySide = new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null);
    return new ExcelCellStyleSnapshot(
        numberFormat,
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new ExcelCellFontSnapshot(
            false,
            false,
            "Aptos",
            ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
            null,
            false,
            false),
        ExcelCellFillSnapshot.pattern(ExcelFillPattern.NONE),
        new ExcelBorderSnapshot(emptySide, emptySide, emptySide, emptySide),
        new ExcelCellProtectionSnapshot(true, false));
  }
}
