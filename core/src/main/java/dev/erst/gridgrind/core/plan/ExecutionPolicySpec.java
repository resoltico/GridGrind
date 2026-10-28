package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import java.util.Objects;

/**
 * Request-side execution policy surface for execution modes, structured journaling, and explicit
 * calculation handling.
 */
@CatalogType(
    id = "ExecutionPolicySpec",
    summary =
        "Optional request execution policy covering execution.mode, execution.journal, execution.calculation, and execution.assertionMode. Omit execution when the standard FULL_XSSF / SUMMARY / DO_NOT_CALCULATE / FAIL_FAST policy is intended, and omit any nested execution field that should keep its own default.",
    displayOrder = 9)
public record ExecutionPolicySpec(
    @ProtocolField(optional = true)
        @WireDefault(factory = ExecutionModeSpec.class, method = "defaults", omit = true)
        ExecutionModeSpec mode,
    @ProtocolField(optional = true)
        @WireDefault(factory = ExecutionJournalSpec.class, method = "defaults", omit = true)
        ExecutionJournalSpec journal,
    @ProtocolField(optional = true)
        @WireDefault(factory = CalculationPolicySpec.class, method = "defaults", omit = true)
        CalculationPolicySpec calculation,
    @ProtocolField(optional = true)
        @WireDefault(factory = AssertionModeSpec.class, method = "defaults", omit = true)
        AssertionModeSpec assertionMode) {
  /** Returns the default execution policy for mode, journaling, calculation, and assertions. */
  public static ExecutionPolicySpec defaults() {
    return new ExecutionPolicySpec(
        ExecutionModeSpec.defaults(),
        ExecutionJournalSpec.defaults(),
        CalculationPolicySpec.defaults(),
        AssertionModeSpec.defaults());
  }

  /** Returns one execution policy that only customizes calculation handling. */
  public static ExecutionPolicySpec calculation(CalculationPolicySpec calculation) {
    return new ExecutionPolicySpec(
        ExecutionModeSpec.defaults(),
        ExecutionJournalSpec.defaults(),
        calculation,
        AssertionModeSpec.defaults());
  }

  /** Returns an execution policy that only sets the execution mode. */
  public static ExecutionPolicySpec mode(ExecutionModeSpec mode) {
    return new ExecutionPolicySpec(
        mode,
        ExecutionJournalSpec.defaults(),
        CalculationPolicySpec.defaults(),
        AssertionModeSpec.defaults());
  }

  /** Returns an execution policy that only customizes journal rendering. */
  public static ExecutionPolicySpec journal(ExecutionJournalSpec journal) {
    return new ExecutionPolicySpec(
        ExecutionModeSpec.defaults(),
        journal,
        CalculationPolicySpec.defaults(),
        AssertionModeSpec.defaults());
  }

  /**
   * Returns an execution policy that customizes mode and journal while leaving calculation at the
   * default.
   */
  public static ExecutionPolicySpec modeAndJournal(
      ExecutionModeSpec mode, ExecutionJournalSpec journal) {
    return new ExecutionPolicySpec(
        mode, journal, CalculationPolicySpec.defaults(), AssertionModeSpec.defaults());
  }

  /**
   * Returns an execution policy that customizes mode and calculation while leaving journaling at
   * the default.
   */
  public static ExecutionPolicySpec modeAndCalculation(
      ExecutionModeSpec mode, CalculationPolicySpec calculation) {
    return new ExecutionPolicySpec(
        mode, ExecutionJournalSpec.defaults(), calculation, AssertionModeSpec.defaults());
  }

  /** Returns an execution policy that collects every terminal-phase assertion outcome. */
  public static ExecutionPolicySpec assertionMode(AssertionModeSpec assertionMode) {
    return new ExecutionPolicySpec(
        ExecutionModeSpec.defaults(),
        ExecutionJournalSpec.defaults(),
        CalculationPolicySpec.defaults(),
        assertionMode);
  }

  public ExecutionPolicySpec {
    Objects.requireNonNull(mode, "mode must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    Objects.requireNonNull(calculation, "calculation must not be null");
    Objects.requireNonNull(assertionMode, "assertionMode must not be null");
  }

  /**
   * Returns whether execution mode, journal, and calculation settings normalize to the defaults.
   */
  public boolean isDefault() {
    return mode.isDefault()
        && journal.isDefault()
        && calculation.isDefault()
        && assertionMode.isDefault();
  }

  /** Returns the effective execution mode after applying GridGrind defaults. */
  public ExecutionModeSpec effectiveMode() {
    return mode;
  }

  /** Returns the effective journal level after applying GridGrind defaults. */
  public ExecutionJournalLevel effectiveJournalLevel() {
    return journal.level();
  }

  /** Returns the effective calculation policy after applying GridGrind defaults. */
  public CalculationPolicySpec effectiveCalculation() {
    return calculation;
  }

  /** Returns the effective assertion mode after applying GridGrind defaults. */
  public AssertionModeSpec effectiveAssertionMode() {
    return assertionMode;
  }
}
