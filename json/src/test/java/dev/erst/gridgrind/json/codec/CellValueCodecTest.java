package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Verifies the shared scalar wire shape and distinct observed/authored number contracts. */
class CellValueCodecTest {
  @Test
  void everyEffectiveScalarUsesTheSameValueField() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    var values =
        List.of(
            new CellValue.Blank(),
            new CellValue.Text("raw"),
            new CellValue.NumberValue(42),
            new CellValue.BooleanValue(true),
            new CellValue.ErrorValue("#CIRCULAR_REF!"));
    var expected =
        List.of(
            "{\"type\":\"BLANK\"}",
            "{\"type\":\"TEXT\",\"value\":\"raw\"}",
            "{\"type\":\"NUMBER\",\"value\":42.0}",
            "{\"type\":\"BOOLEAN\",\"value\":true}",
            "{\"type\":\"ERROR\",\"value\":\"#CIRCULAR_REF!\"}");
    for (int index = 0; index < values.size(); index++) {
      byte[] bytes = mapper.writeValueAsBytes(values.get(index));
      assertEquals(mapper.readTree(expected.get(index)), mapper.readTree(bytes));
      assertEquals(values.get(index), mapper.readValue(bytes, CellValue.class));
    }
    assertEquals(
        mapper.readTree("{\"type\":\"TEXT\",\"value\":{\"type\":\"INLINE\",\"text\":\"raw\"}}"),
        mapper.valueToTree(new CellValueSpec.Text(TextSourceSpec.inline("raw"))));
  }

  @ParameterizedTest
  @ValueSource(doubles = {Double.NaN, Double.POSITIVE_INFINITY, Double.NEGATIVE_INFINITY})
  void observedNonfiniteNumbersRoundTripWithoutBecomingAuthorableExpectations(double value) {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    var report =
        new CellReport.NumberReport(
            new CellAddress("A1"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new CellValue.NumberValue(value)),
            Optional.empty());
    byte[] bytes = mapper.writeValueAsBytes(report);
    assertTrue(mapper.readTree(bytes).path("value").path("value").isString());
    assertEquals(report, mapper.readValue(bytes, CellReport.class));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.NumberValue(value));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellAssertion.Value(new CellValue.NumberValue(value)));
  }
}
