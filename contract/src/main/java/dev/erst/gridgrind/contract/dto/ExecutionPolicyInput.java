package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/**
 * Request-side execution policy surface for low-memory execution modes, structured journaling, and
 * explicit calculation handling.
 */
public record ExecutionPolicyInput(
    @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = ExecutionModeInput.DefaultFilter.class)
        ExecutionModeInput mode,
    @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = ExecutionJournalInput.DefaultFilter.class)
        ExecutionJournalInput journal,
    @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = CalculationPolicyInput.DefaultFilter.class)
        CalculationPolicyInput calculation) {
  /** Returns the default execution policy for mode, journaling, and calculation handling. */
  public static ExecutionPolicyInput defaults() {
    return new ExecutionPolicyInput(
        ExecutionModeInput.defaults(),
        ExecutionJournalInput.defaults(),
        CalculationPolicyInput.defaults());
  }

  /** Returns one execution policy that only customizes calculation handling. */
  public static ExecutionPolicyInput calculation(CalculationPolicyInput calculation) {
    return new ExecutionPolicyInput(
        ExecutionModeInput.defaults(), ExecutionJournalInput.defaults(), calculation);
  }

  /** Returns an execution policy that only sets the execution mode. */
  public static ExecutionPolicyInput mode(ExecutionModeInput mode) {
    return new ExecutionPolicyInput(
        mode, ExecutionJournalInput.defaults(), CalculationPolicyInput.defaults());
  }

  /** Returns an execution policy that only customizes journal rendering. */
  public static ExecutionPolicyInput journal(ExecutionJournalInput journal) {
    return new ExecutionPolicyInput(
        ExecutionModeInput.defaults(), journal, CalculationPolicyInput.defaults());
  }

  /**
   * Returns an execution policy that customizes mode and journal while leaving calculation at the
   * default.
   */
  public static ExecutionPolicyInput modeAndJournal(
      ExecutionModeInput mode, ExecutionJournalInput journal) {
    return new ExecutionPolicyInput(mode, journal, CalculationPolicyInput.defaults());
  }

  /**
   * Returns an execution policy that customizes mode and calculation while leaving journaling at
   * the default.
   */
  public static ExecutionPolicyInput modeAndCalculation(
      ExecutionModeInput mode, CalculationPolicyInput calculation) {
    return new ExecutionPolicyInput(mode, ExecutionJournalInput.defaults(), calculation);
  }

  /** Reads one execution-policy block while applying the documented omission defaults. */
  @JsonCreator
  static ExecutionPolicyInput create(
      @JsonProperty("mode") ExecutionModeInput mode,
      @JsonProperty("journal") ExecutionJournalInput journal,
      @JsonProperty("calculation") CalculationPolicyInput calculation) {
    return new ExecutionPolicyInput(
        mode == null ? ExecutionModeInput.defaults() : mode,
        journal == null ? ExecutionJournalInput.defaults() : journal,
        calculation == null ? CalculationPolicyInput.defaults() : calculation);
  }

  public ExecutionPolicyInput {
    Objects.requireNonNull(mode, "mode must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    Objects.requireNonNull(calculation, "calculation must not be null");
  }

  /**
   * Returns whether execution mode, journal, and calculation settings normalize to the defaults.
   */
  @JsonIgnore
  public boolean isDefault() {
    return mode.isDefault() && journal.isDefault() && calculation.isDefault();
  }

  /** Returns the effective execution mode after applying GridGrind defaults. */
  @JsonIgnore
  public ExecutionModeInput effectiveMode() {
    return mode;
  }

  /** Returns the effective journal level after applying GridGrind defaults. */
  @JsonIgnore
  public ExecutionJournalLevel effectiveJournalLevel() {
    return journal.level();
  }

  /** Returns the effective calculation policy after applying GridGrind defaults. */
  @JsonIgnore
  public CalculationPolicyInput effectiveCalculation() {
    return calculation;
  }

  /** Custom Jackson inclusion filter that omits the standard default execution policy. */
  public static final class DefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null || (other instanceof ExecutionPolicyInput input && input.isDefault());
    }

    @Override
    public int hashCode() {
      return DefaultFilter.class.hashCode();
    }
  }
}
