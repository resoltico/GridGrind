package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.formula.CalculationPolicyInput;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import java.util.Objects;

/**
 * Request-side execution policy surface for low-memory execution modes, structured journaling, and
 * explicit calculation handling.
 */
@CatalogType(
    id = "ExecutionPolicyInput",
    summary =
        "Optional request execution policy covering execution.mode, execution.journal, execution.calculation, and execution.assertionMode. Omit execution when the standard FULL_XSSF / SUMMARY / DO_NOT_CALCULATE / FAIL_FAST policy is intended, and omit any nested execution field that should keep its own default.",
    displayOrder = 9)
public record ExecutionPolicyInput(
    @ProtocolField(optional = true)
        @WireDefault(factory = ExecutionModeInput.class, method = "defaults", omit = true)
        ExecutionModeInput mode,
    @ProtocolField(optional = true)
        @WireDefault(factory = ExecutionJournalInput.class, method = "defaults", omit = true)
        ExecutionJournalInput journal,
    @ProtocolField(optional = true)
        @WireDefault(factory = CalculationPolicyInput.class, method = "defaults", omit = true)
        CalculationPolicyInput calculation,
    @ProtocolField(optional = true)
        @WireDefault(factory = AssertionModeInput.class, method = "defaults", omit = true)
        AssertionModeInput assertionMode) {
  /** Returns the default execution policy for mode, journaling, calculation, and assertions. */
  public static ExecutionPolicyInput defaults() {
    return new ExecutionPolicyInput(
        ExecutionModeInput.defaults(),
        ExecutionJournalInput.defaults(),
        CalculationPolicyInput.defaults(),
        AssertionModeInput.defaults());
  }

  /** Returns one execution policy that only customizes calculation handling. */
  public static ExecutionPolicyInput calculation(CalculationPolicyInput calculation) {
    return new ExecutionPolicyInput(
        ExecutionModeInput.defaults(),
        ExecutionJournalInput.defaults(),
        calculation,
        AssertionModeInput.defaults());
  }

  /** Returns an execution policy that only sets the execution mode. */
  public static ExecutionPolicyInput mode(ExecutionModeInput mode) {
    return new ExecutionPolicyInput(
        mode,
        ExecutionJournalInput.defaults(),
        CalculationPolicyInput.defaults(),
        AssertionModeInput.defaults());
  }

  /** Returns an execution policy that only customizes journal rendering. */
  public static ExecutionPolicyInput journal(ExecutionJournalInput journal) {
    return new ExecutionPolicyInput(
        ExecutionModeInput.defaults(),
        journal,
        CalculationPolicyInput.defaults(),
        AssertionModeInput.defaults());
  }

  /**
   * Returns an execution policy that customizes mode and journal while leaving calculation at the
   * default.
   */
  public static ExecutionPolicyInput modeAndJournal(
      ExecutionModeInput mode, ExecutionJournalInput journal) {
    return new ExecutionPolicyInput(
        mode, journal, CalculationPolicyInput.defaults(), AssertionModeInput.defaults());
  }

  /**
   * Returns an execution policy that customizes mode and calculation while leaving journaling at
   * the default.
   */
  public static ExecutionPolicyInput modeAndCalculation(
      ExecutionModeInput mode, CalculationPolicyInput calculation) {
    return new ExecutionPolicyInput(
        mode, ExecutionJournalInput.defaults(), calculation, AssertionModeInput.defaults());
  }

  /** Returns an execution policy that collects every terminal-phase assertion outcome. */
  public static ExecutionPolicyInput assertionMode(AssertionModeInput assertionMode) {
    return new ExecutionPolicyInput(
        ExecutionModeInput.defaults(),
        ExecutionJournalInput.defaults(),
        CalculationPolicyInput.defaults(),
        assertionMode);
  }

  public ExecutionPolicyInput {
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
  public ExecutionModeInput effectiveMode() {
    return mode;
  }

  /** Returns the effective journal level after applying GridGrind defaults. */
  public ExecutionJournalLevel effectiveJournalLevel() {
    return journal.level();
  }

  /** Returns the effective calculation policy after applying GridGrind defaults. */
  public CalculationPolicyInput effectiveCalculation() {
    return calculation;
  }

  /** Returns the effective assertion mode after applying GridGrind defaults. */
  public AssertionModeInput effectiveAssertionMode() {
    return assertionMode;
  }
}
