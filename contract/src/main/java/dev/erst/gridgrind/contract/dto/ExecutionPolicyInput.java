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
            valueFilter = ExecutionPolicyInput.ExecutionModeDefaultFilter.class)
        ExecutionModeInput mode,
    @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = ExecutionPolicyInput.ExecutionJournalDefaultFilter.class)
        ExecutionJournalInput journal,
    @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = ExecutionPolicyInput.CalculationPolicyDefaultFilter.class)
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

  public ExecutionPolicyInput {
    Objects.requireNonNull(mode, "mode must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    Objects.requireNonNull(calculation, "calculation must not be null");
  }

  /** Creates an execution policy that only sets the read/write execution mode family. */
  public ExecutionPolicyInput(ExecutionModeInput mode) {
    this(mode, ExecutionJournalInput.defaults(), CalculationPolicyInput.defaults());
  }

  /** Creates an execution policy that only customizes journal rendering. */
  public ExecutionPolicyInput(ExecutionJournalInput journal) {
    this(ExecutionModeInput.defaults(), journal, CalculationPolicyInput.defaults());
  }

  /**
   * Creates an execution policy that sets mode and journal settings but leaves calculation
   * defaulted.
   */
  public ExecutionPolicyInput(ExecutionModeInput mode, ExecutionJournalInput journal) {
    this(mode, journal, CalculationPolicyInput.defaults());
  }

  /**
   * Creates an execution policy that sets mode and calculation while leaving journaling defaulted.
   */
  public ExecutionPolicyInput(ExecutionModeInput mode, CalculationPolicyInput calculation) {
    this(mode, ExecutionJournalInput.defaults(), calculation);
  }

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

  /**
   * Returns whether execution mode, journal, and calculation settings normalize to the defaults.
   */
  @JsonIgnore
  public boolean isDefault() {
    return mode.isDefault() && journal.isDefault() && calculation.isDefault();
  }

  /** Custom Jackson inclusion filter that omits the default execution-policy object. */
  public static final class DefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null || (other instanceof ExecutionPolicyInput input && input.isDefault());
    }

    @Override
    public int hashCode() {
      return 0;
    }
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

  /** Custom Jackson inclusion filter that omits the default execution mode object. */
  public static final class ExecutionModeDefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null || (other instanceof ExecutionModeInput input && input.isDefault());
    }

    @Override
    public int hashCode() {
      return 0;
    }
  }

  /** Custom Jackson inclusion filter that omits the default execution journal object. */
  public static final class ExecutionJournalDefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null || (other instanceof ExecutionJournalInput input && input.isDefault());
    }

    @Override
    public int hashCode() {
      return 0;
    }
  }

  /** Custom Jackson inclusion filter that omits the default calculation policy object. */
  public static final class CalculationPolicyDefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null || (other instanceof CalculationPolicyInput input && input.isDefault());
    }

    @Override
    public int hashCode() {
      return 0;
    }
  }
}
