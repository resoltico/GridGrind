package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestMissingTypeDiscriminator;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import dev.erst.gridgrind.core.problem.RequestUnknownTypeDiscriminator;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.ValueSource;

/** Selector is an unannotated marker; its family declarations own the shared type discriminator. */
class SelectorDiscriminatorIntakeTest {
  @ParameterizedTest
  @MethodSource("selectorCases")
  void everySelectorFamilyDecodesThroughItsDeclaredDiscriminator(SelectorCase sample) {
    byte[] bytes = ("{\"target\":" + sample.json() + "}").getBytes(StandardCharsets.UTF_8);
    var parsed = TolerantRequestJsonParser.parse(bytes);
    assertTrue(parsed.problems().isEmpty());
    List<RequestStructuralProblem> problems = new ArrayList<>();
    RequestSchemaStructure.validateRecord(parsed.root(), SelectorHolder.class, "", 0, problems);
    assertTrue(problems.isEmpty(), () -> sample.family().getSimpleName() + ": " + problems);
    List<RequestBindingFailure> failures = new ArrayList<>();
    var actual =
        RequestFragmentBinder.bindNode(parsed.root(), SelectorHolder.class, "", problems, failures);
    assertTrue(failures.isEmpty(), failures::toString);
    assertEquals(sample.concrete(), actual.orElseThrow().target().getClass());
  }

  @Test
  void selectorCasesCoverEveryDeclaredFamily() {
    assertEquals(
        Set.of(Selector.class.getPermittedSubclasses()),
        selectorCases().map(SelectorCase::family).collect(Collectors.toSet()));
  }

  @Test
  void canonicalRootAcceptsWorkbookAndRangeTargetsWithoutAddingATagToStructuralSteps() {
    var intake =
        GridGrindJson.analyzeRequest(
            request(
                "{\"stepId\":\"workbook\",\"target\":{\"type\":\"WORKBOOK_CURRENT\"},\"query\":{\"type\":\"GET_WORKBOOK_SUMMARY\"}},"
                    + "{\"stepId\":\"ranges\",\"target\":{\"type\":\"RANGE_BY_RANGE\",\"sheetName\":\"Budget\",\"range\":\"A1:B2\"},\"query\":{\"type\":\"GET_DATA_VALIDATIONS\"}}"));
    assertTrue(intake.isBindable(), intake.structuralProblems()::toString);
    assertEquals(
        WorkbookSelector.Current.class,
        intake.requireCompletePlan().steps().get(0).target().getClass());
    assertEquals(
        RangeSelector.ByRange.class,
        intake.requireCompletePlan().steps().get(1).target().getClass());
  }

  @ParameterizedTest
  @ValueSource(strings = {"{}", "{\"type\":\"UNKNOWN_SELECTOR\"}"})
  void missingAndUnknownSelectorTagsRetainTheValidSiblingAndExactUnknownTokenOffset(
      String selector) {
    byte[] bytes =
        request(
            "{\"stepId\":\"valid\",\"target\":{\"type\":\"WORKBOOK_CURRENT\"},\"query\":{\"type\":\"GET_WORKBOOK_SUMMARY\"}},"
                + "{\"stepId\":\"invalid\",\"target\":"
                + selector
                + ",\"query\":{\"type\":\"GET_WORKBOOK_SUMMARY\"}}");
    var intake = GridGrindJson.analyzeRequest(bytes);
    assertFalse(intake.isBindable());
    var steps = intake.boundFragments().steps().orElseThrow();
    assertTrue(steps.get(0).value().isPresent());
    assertTrue(steps.get(1).value().isEmpty());
    if ("{}".equals(selector)) {
      assertTrue(
          intake.structuralProblems().stream()
              .anyMatch(problem ->
                  problem instanceof RequestMissingTypeDiscriminator
                      && problem.jsonPath().equals(Optional.of("steps[1].target.type"))));
    } else {
      var unknown =
          intake.structuralProblems().stream()
              .filter(RequestUnknownTypeDiscriminator.class::isInstance)
              .map(RequestUnknownTypeDiscriminator.class::cast)
              .findFirst()
              .orElseThrow();
      assertEquals(Optional.of("steps[1].target.type"), unknown.jsonPath());
      assertEquals(
          Optional.of(
              (long)
                  new String(bytes, StandardCharsets.UTF_8)
                      .indexOf("\"type\":\"UNKNOWN_SELECTOR\"")),
          unknown.byteOffset());
    }
  }

  @Test
  void inconsistentFamilyDiscriminatorsFailAsAModelDefect() {
    assertThrows(
        IllegalStateException.class,
        () ->
            dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport.discriminatorField(
                Mixed.class));
  }

  /** Test marker whose family tags conflict, exercising model-defect rejection. */
  sealed interface Mixed permits TypeFamily, StatusFamily {}

  /** Test family declaring the type tag. */
  @dev.erst.gridgrind.core.metadata.Union(discriminator = "type")
  non-sealed interface TypeFamily extends Mixed {}

  /** Test family declaring the status tag. */
  @dev.erst.gridgrind.core.metadata.Union(discriminator = "status")
  non-sealed interface StatusFamily extends Mixed {}

  private static byte[] request(String steps) {
    return ("{\"protocolVersion\":1,\"source\":{\"type\":\"NEW\"},\"persistence\":{\"type\":\"NONE\"},\"steps\":["
            + steps
            + "]}")
        .getBytes(StandardCharsets.UTF_8);
  }

  static Stream<SelectorCase> selectorCases() {
    return Stream.of(
        new SelectorCase(
            WorkbookSelector.class,
            WorkbookSelector.Current.class,
            "{\"type\":\"WORKBOOK_CURRENT\"}"),
        new SelectorCase(
            SheetSelector.class,
            SheetSelector.ByName.class,
            "{\"type\":\"SHEET_BY_NAME\",\"name\":\"Budget\"}"),
        new SelectorCase(
            CellSelector.class,
            CellSelector.ByAddress.class,
            "{\"type\":\"CELL_BY_ADDRESS\",\"sheetName\":\"Budget\",\"address\":\"A1\"}"),
        new SelectorCase(
            RangeSelector.class,
            RangeSelector.ByRange.class,
            "{\"type\":\"RANGE_BY_RANGE\",\"sheetName\":\"Budget\",\"range\":\"A1:B2\"}"),
        new SelectorCase(
            RowBandSelector.class,
            RowBandSelector.Span.class,
            "{\"type\":\"ROW_BAND_SPAN\",\"sheetName\":\"Budget\",\"firstRowIndex\":0,\"lastRowIndex\":1}"),
        new SelectorCase(
            ColumnBandSelector.class,
            ColumnBandSelector.Span.class,
            "{\"type\":\"COLUMN_BAND_SPAN\",\"sheetName\":\"Budget\",\"firstColumnIndex\":0,\"lastColumnIndex\":1}"),
        new SelectorCase(
            DrawingObjectSelector.class,
            DrawingObjectSelector.ByName.class,
            "{\"type\":\"DRAWING_OBJECT_BY_NAME\",\"sheetName\":\"Budget\",\"objectName\":\"Logo\"}"),
        new SelectorCase(
            ChartSelector.class,
            ChartSelector.ByName.class,
            "{\"type\":\"CHART_BY_NAME\",\"sheetName\":\"Budget\",\"chartName\":\"Chart1\"}"),
        new SelectorCase(
            TableSelector.class,
            TableSelector.ByName.class,
            "{\"type\":\"TABLE_BY_NAME\",\"name\":\"Orders\"}"),
        new SelectorCase(
            PivotTableSelector.class,
            PivotTableSelector.ByName.class,
            "{\"type\":\"PIVOT_TABLE_BY_NAME\",\"name\":\"Totals\"}"),
        new SelectorCase(
            NamedRangeSelector.class,
            NamedRangeSelector.ByName.class,
            "{\"type\":\"NAMED_RANGE_BY_NAME\",\"name\":\"Revenue\"}"),
        new SelectorCase(
            TableRowSelector.class,
            TableRowSelector.ByIndex.class,
            "{\"type\":\"TABLE_ROW_BY_INDEX\",\"table\":{\"type\":\"TABLE_BY_NAME\",\"name\":\"Orders\"},\"rowIndex\":0}"),
        new SelectorCase(
            TableCellSelector.class,
            TableCellSelector.ByColumnName.class,
            "{\"type\":\"TABLE_CELL_BY_COLUMN_NAME\",\"row\":{\"type\":\"TABLE_ROW_BY_INDEX\",\"table\":{\"type\":\"TABLE_BY_NAME\",\"name\":\"Orders\"},\"rowIndex\":0},\"columnName\":\"Amount\"}"));
  }

  record SelectorCase(
      Class<? extends Selector> family, Class<? extends Selector> concrete, String json) {}

  public record SelectorHolder(Selector target) {}
}
