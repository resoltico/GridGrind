package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.ExcelGradientFillGeometry;
import org.junit.jupiter.api.Test;

/** Tests the shared authored-gradient geometry rules. */
class ExcelGradientFillGeometryTest {
  @Test
  void effectiveTypeInfersPathForEachIndividualOffset() {
    assertEquals("PATH", ExcelGradientFillGeometry.effectiveType(null, 0.1d, null, null, null));
    assertEquals("PATH", ExcelGradientFillGeometry.effectiveType(null, null, 0.2d, null, null));
    assertEquals("PATH", ExcelGradientFillGeometry.effectiveType(null, null, null, 0.3d, null));
    assertEquals("PATH", ExcelGradientFillGeometry.effectiveType(null, null, null, null, 0.4d));
  }

  @Test
  void effectiveTypeDefaultsToLinearWhenNoOffsetsExist() {
    assertEquals("LINEAR", ExcelGradientFillGeometry.effectiveType(null, null, null, null, null));
  }

  @Test
  void effectiveTypeHonorsExplicitTypeAfterNormalization() {
    assertEquals("PATH", ExcelGradientFillGeometry.effectiveType(" path ", null, null, null, null));
    assertEquals(
        "LINEAR", ExcelGradientFillGeometry.effectiveType(" linear ", 0.1d, null, null, null));
  }

  @Test
  void requireCompatibleGeometryRejectsMixedModels() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelGradientFillGeometry.requireCompatibleGeometry(
                "LINEAR", 12.0d, 0.1d, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelGradientFillGeometry.requireCompatibleGeometry(
                "PATH", 12.0d, null, null, null, null));
  }
}
