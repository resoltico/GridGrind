package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.formula.CalculationPolicyInput;
import dev.erst.gridgrind.core.model.formula.CalculationStrategyInput;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.plan.AssertionModeInput;
import dev.erst.gridgrind.core.plan.ExecutionJournalInput;
import dev.erst.gridgrind.core.plan.ExecutionModeInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
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
        new ExecutionJournalInput(ExecutionJournalLevel.SUMMARY),
        mapper.readValue("{}", ExecutionJournalInput.class));
    assertEquals("{}", mapper.writeValueAsString(ExecutionJournalInput.defaults()));
    assertEquals(
        new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE),
        mapper.readValue("{\"level\":\"VERBOSE\"}", ExecutionJournalInput.class));
  }

  @Test
  void nestedPolicyDefaultsPreserveExplicitOverrides() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    ExecutionPolicyInput defaults =
        new ExecutionPolicyInput(
            new ExecutionModeInput.FullXssf(),
            new ExecutionJournalInput(ExecutionJournalLevel.SUMMARY),
            new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), false),
            AssertionModeInput.FAIL_FAST);
    assertEquals(defaults, mapper.readValue("{}", ExecutionPolicyInput.class));
    assertEquals("{}", mapper.writeValueAsString(defaults));
    assertEquals(
        new ExecutionPolicyInput(
            defaults.mode(),
            new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE),
            defaults.calculation(),
            AssertionModeInput.COLLECT),
        mapper.readValue(
            "{\"journal\":{\"level\":\"VERBOSE\"},\"assertionMode\":\"COLLECT\"}",
            ExecutionPolicyInput.class));
    assertEquals(
        new FormulaEnvironmentInput(List.of(), FormulaMissingWorkbookPolicy.ERROR, List.of()),
        mapper.readValue("{}", FormulaEnvironmentInput.class));
    assertThrows(
        JacksonException.class,
        () -> mapper.readValue("{\"journal\":null}", ExecutionPolicyInput.class));
  }
}
