package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.core.vocabulary.WireString;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.SplittableRandom;
import java.util.function.Function;
import java.util.stream.Stream;
import org.apache.poi.ss.util.CellReference;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestFactory;
import tools.jackson.core.JacksonException;
import tools.jackson.core.type.TypeReference;

/** Exercises generated identifier values against constructor rules and the actual scalar codec. */
class WireStringPropertyTest {
  @TestFactory
  @SuppressWarnings("PMD.UnitTestShouldIncludeAssert") // Assertions execute in each dynamic case.
  Stream<DynamicTest> validAndInvalidDomainStrings() {
    return domains().stream()
        .flatMap(
            domain ->
                Stream.of(
                    DynamicTest.dynamicTest(
                        domain.type().getSimpleName() + "/generated-valid",
                        () -> {
                          for (String text : domain.valid()) {
                            WireString value = domain.construct().apply(text);
                            byte[] bytes =
                                GridGrindJsonMapperSupport.JSON_MAPPER.writeValueAsBytes(value);
                            assertEquals(
                                value,
                                GridGrindJsonMapperSupport.JSON_MAPPER.readValue(
                                    bytes, domain.type()),
                                text);
                            String expected =
                                domain.type() == RgbHex.class
                                    ? text.toUpperCase(java.util.Locale.ROOT)
                                    : text;
                            assertEquals(expected, value.value());
                            assertArrayEquals(
                                GridGrindJsonMapperSupport.JSON_MAPPER.writeValueAsBytes(expected),
                                bytes);
                            assertEquals(value.value(), value.toString());
                          }
                        }),
                    DynamicTest.dynamicTest(
                        domain.type().getSimpleName() + "/generated-invalid",
                        () -> {
                          assertThrows(
                              NullPointerException.class, () -> domain.construct().apply(null));
                          for (String text : domain.invalid()) {
                            assertThrows(
                                IllegalArgumentException.class,
                                () -> domain.construct().apply(text),
                                text);
                            String json =
                                GridGrindJsonMapperSupport.JSON_MAPPER.writeValueAsString(text);
                            assertThrows(
                                InvalidRequestException.class,
                                () ->
                                    GridGrindJsonMapperSupport.JSON_MAPPER.readValue(
                                        json, domain.type()),
                                text);
                          }
                          for (String json : List.of("1", "true", "{}", "[]")) {
                            assertThrows(
                                JacksonException.class,
                                () ->
                                    GridGrindJsonMapperSupport.JSON_MAPPER.readValue(
                                        json, domain.type()),
                                json);
                          }
                        })));
  }

  @Test
  void oversizedRowsReportTheDomainBoundInsteadOfAnIntegerParseFailure() {
    String address = "A" + "9".repeat(128);
    assertEquals(
        Optional.of(
            dev.erst.gridgrind.core.vocabulary.ProtocolCellAddressValidation.Violation.BOUNDS),
        dev.erst.gridgrind.core.vocabulary.ProtocolCellAddressValidation.violation(address));
    assertEquals(
        "address must be within Excel .xlsx bounds",
        assertThrows(IllegalArgumentException.class, () -> new CellAddress(address)).getMessage());
  }

  @Test
  void optionalAndCollectionValuesRetainTheirStringShape() {
    Envelope value =
        new Envelope(
            Optional.of(new SheetName("Budget")),
            List.of(new CellAddress("A1"), new CellAddress("$XFD$1048576")));
    String json = GridGrindJsonMapperSupport.WIRE_JSON_MAPPER.writeValueAsString(value);
    assertEquals("{\"sheet\":\"Budget\",\"cells\":[\"A1\",\"$XFD$1048576\"]}", json);
    assertEquals(value, GridGrindJsonMapperSupport.JSON_MAPPER.readValue(json, Envelope.class));
    String missing =
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER.writeValueAsString(
            new Envelope(Optional.empty(), List.of()));
    assertEquals("{\"cells\":[]}", missing);
    assertEquals(
        Optional.empty(),
        GridGrindJsonMapperSupport.JSON_MAPPER.readValue(missing, Envelope.class).sheet());
    List<SheetName> sheets =
        GridGrindJsonMapperSupport.JSON_MAPPER.readValue(
            "[\"Budget\",\"Forecast\"]", new TypeReference<List<SheetName>>() {});
    assertEquals(List.of(new SheetName("Budget"), new SheetName("Forecast")), sheets);
    assertEquals("B2:A1", new RangeRef("B2:A1").value());
  }

  private static List<Domain> domains() {
    SplittableRandom random = new SplittableRandom(20261025);
    List<String> cells = new ArrayList<>(List.of("A1", "$a$1", "XFD1048576"));
    List<String> ranges = new ArrayList<>(List.of("A1", "$A1:B$2", "B2:A1"));
    List<String> sheets = new ArrayList<>(List.of("Budget", "予算", "a".repeat(31), "Forecast 2026"));
    List<String> names = new ArrayList<>(List.of("Revenue_Total", "_budget", "売上", "\\name"));
    List<String> colors = new ArrayList<>(List.of("#000000", "#FFFFFF", "#a1b2c3"));
    List<String> steps = new ArrayList<>(List.of("step-1", "a.b_c-2"));
    List<String> plans = new ArrayList<>(List.of("plan-1", "Yearly plan 2026"));
    List<String> invalidSheets =
        new ArrayList<>(
            List.of(
                "",
                " ",
                "n".repeat(32),
                "bad/name",
                "bad:name",
                "[name]",
                "'name",
                "name'",
                "a\u0000b"));
    List<String> invalidCells =
        new ArrayList<>(
            List.of(
                "", " ", "A0", "A1048577", "XFE1", "A1:B2", "Budget!A1", "A999999999999999999999"));
    List<String> invalidRanges =
        new ArrayList<>(
            List.of("", " ", "A1:", ":B2", "A1:B2:C3", "XFE1:A1", "A0:B2", "Budget!A1:B2"));
    List<String> invalidNames =
        new ArrayList<>(
            List.of("", " ", "1name", "bad name", "_xlnm.Test", "A1", "R1C1", "n".repeat(256)));
    List<String> invalidColors =
        new ArrayList<>(List.of("", " ", "000000", "#000", "#1234567", "#G00000"));
    List<String> invalidSteps = new ArrayList<>(List.of("", " ", "step 1", "step/1", "é"));
    List<String> invalidPlans = new ArrayList<>(List.of("", " ", "\t\n"));
    for (int index = 0; index < 128; index++) {
      String cell =
          CellReference.convertNumToColString(random.nextInt(16384)) + random.nextInt(1, 1048577);
      cells.add(cell);
      ranges.add(
          cell
              + ":"
              + CellReference.convertNumToColString(random.nextInt(16384))
              + random.nextInt(1, 1048577));
      sheets.add("Sheet " + index);
      names.add("Name_" + index);
      colors.add("#%06x".formatted(random.nextInt(0x1000000)));
      steps.add("step_" + index);
      plans.add("plan " + index);
      invalidSheets.add("Sheet" + index + "/");
      invalidSheets.add("s".repeat(32 + index));
      invalidCells.add(CellReference.convertNumToColString(16384 + index) + "1");
      invalidCells.add("A" + (1048577 + index));
      invalidRanges.add(cell + ":XFE" + (index + 1));
      invalidNames.add("1Name_" + index);
      invalidColors.add("#G%05x".formatted(random.nextInt(0x100000)));
      invalidSteps.add("step " + index);
      invalidPlans.add(" ".repeat(index + 1));
    }
    return List.of(
        new Domain(SheetName.class, SheetName::new, sheets, invalidSheets),
        new Domain(
            SheetLabel.class,
            SheetLabel::new,
            List.of("Budget", "invalid[]", "s".repeat(100), "予算"),
            invalidPlans),
        new Domain(CellAddress.class, CellAddress::new, cells, invalidCells),
        new Domain(RangeRef.class, RangeRef::new, ranges, invalidRanges),
        new Domain(DefinedName.class, DefinedName::new, names, invalidNames),
        new Domain(TableName.class, TableName::new, names, invalidNames),
        new Domain(RgbHex.class, RgbHex::new, colors, invalidColors),
        new Domain(StepId.class, StepId::new, steps, invalidSteps),
        new Domain(PlanId.class, PlanId::new, plans, invalidPlans));
  }

  private record Domain(
      Class<? extends WireString> type,
      Function<String, WireString> construct,
      List<String> valid,
      List<String> invalid) {}

  private record Envelope(Optional<SheetName> sheet, List<CellAddress> cells) {}
}
