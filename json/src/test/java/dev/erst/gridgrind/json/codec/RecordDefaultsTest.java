package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.plan.AssertionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;

/** Checks omission defaults independently of creator factories and explicit values. */
class RecordDefaultsTest {
  /** Fixture with independent boolean and list defaults plus an absent optional. */
  public record Settings(
      @WireDefault(value = "true", omit = true) boolean enabled,
      @WireDefault(value = "[]", omit = true) List<String> labels,
      Optional<String> note) {}

  /** Fixture whose first field depends on a later component. */
  public record Named(@WireDefault(from = "source", omit = true) String display, String source) {}

  @Test
  void omittedComponentsUseDefaultsButExplicitValuesSurvive() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    assertEquals(
        new Settings(true, List.of(), Optional.empty()), mapper.readValue("{}", Settings.class));
    assertEquals(
        new Settings(false, List.of("chosen"), Optional.of("note")),
        mapper.readValue(
            "{\"enabled\":false,\"labels\":[\"chosen\"],\"note\":\"note\"}", Settings.class));
    assertEquals("{}", mapper.writeValueAsString(new Settings(true, List.of(), Optional.empty())));
    assertEquals(
        "{\"enabled\":false,\"labels\":[\"chosen\"]}",
        mapper.writeValueAsString(new Settings(false, List.of("chosen"), Optional.empty())));
  }

  @Test
  void explicitNullDoesNotApplyAnOmissionDefault() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    assertThrows(
        JacksonException.class, () -> mapper.readValue("{\"enabled\":null}", Settings.class));
    assertThrows(
        JacksonException.class, () -> mapper.readValue("{\"labels\":null}", Settings.class));
  }

  @Test
  void derivedDefaultsUseTheNamedSiblingRegardlessOfWireOrder() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    assertEquals(
        new Named("Amount", "Amount"), mapper.readValue("{\"source\":\"Amount\"}", Named.class));
    assertEquals(
        new Named("Total", "Amount"),
        mapper.readValue("{\"display\":\"Total\",\"source\":\"Amount\"}", Named.class));
    assertEquals(
        "{\"source\":\"Amount\"}", mapper.writeValueAsString(new Named("Amount", "Amount")));
  }

  @Test
  void journalUsesCanonicalConstructorWithMetadataDefault() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    assertEquals(
        new ExecutionJournalSpec(ExecutionJournalLevel.SUMMARY),
        mapper.readValue("{}", ExecutionJournalSpec.class));
    assertEquals("{}", mapper.writeValueAsString(ExecutionJournalSpec.defaults()));
    assertEquals(
        new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
        mapper.readValue("{\"level\":\"VERBOSE\"}", ExecutionJournalSpec.class));
  }

  @Test
  void nestedPolicyDefaultsPreserveExplicitOverrides() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    ExecutionPolicySpec defaults =
        new ExecutionPolicySpec(
            new ExecutionModeSpec.FullXssf(),
            new ExecutionJournalSpec(ExecutionJournalLevel.SUMMARY),
            new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), false),
            AssertionModeSpec.FAIL_FAST);
    assertEquals(defaults, mapper.readValue("{}", ExecutionPolicySpec.class));
    assertEquals("{}", mapper.writeValueAsString(defaults));
    assertEquals(
        new ExecutionPolicySpec(
            defaults.mode(),
            new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
            defaults.calculation(),
            AssertionModeSpec.COLLECT),
        mapper.readValue(
            "{\"journal\":{\"level\":\"VERBOSE\"},\"assertionMode\":\"COLLECT\"}",
            ExecutionPolicySpec.class));
    assertEquals(
        new FormulaEnvironmentSpec(List.of(), FormulaMissingWorkbookPolicy.ERROR, List.of()),
        mapper.readValue("{}", FormulaEnvironmentSpec.class));
    assertThrows(
        JacksonException.class,
        () -> mapper.readValue("{\"journal\":null}", ExecutionPolicySpec.class));
  }
}
