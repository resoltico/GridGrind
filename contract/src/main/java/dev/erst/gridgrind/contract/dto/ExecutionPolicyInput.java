package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;

/**
 * Request-side execution policy surface for low-memory execution modes, structured journaling, and
 * explicit calculation handling.
 */
public record ExecutionPolicyInput(
    @JsonInclude(JsonInclude.Include.NON_NULL) ExecutionModeInput mode,
    @JsonInclude(JsonInclude.Include.NON_NULL) ExecutionJournalInput journal,
    @JsonInclude(JsonInclude.Include.NON_NULL) CalculationPolicyInput calculation) {
  public ExecutionPolicyInput {
    mode = mode == null || mode.isDefault() ? null : mode;
    journal = journal == null || journal.isDefault() ? null : journal;
    calculation = calculation == null || calculation.isDefault() ? null : calculation;
  }

  /** Creates an execution policy that only sets the read/write execution mode family. */
  public ExecutionPolicyInput(ExecutionModeInput mode) {
    this(mode, null, null);
  }

  /**
   * Creates an execution policy that sets mode and journal settings but leaves calculation
   * defaulted.
   */
  public ExecutionPolicyInput(ExecutionModeInput mode, ExecutionJournalInput journal) {
    this(mode, journal, null);
  }

  /**
   * Returns whether execution mode, journal, and calculation settings normalize to the defaults.
   */
  @JsonIgnore
  public boolean isDefault() {
    return mode == null && journal == null && calculation == null;
  }

  /** Returns the effective execution mode after applying GridGrind defaults. */
  @JsonIgnore
  public ExecutionModeInput effectiveMode() {
    return mode == null ? new ExecutionModeInput(null, null) : mode;
  }

  /** Returns the effective journal level after applying GridGrind defaults. */
  @JsonIgnore
  public ExecutionJournalLevel effectiveJournalLevel() {
    return ExecutionJournalInput.effectiveLevel(journal);
  }

  /** Returns the effective calculation policy after applying GridGrind defaults. */
  @JsonIgnore
  public CalculationPolicyInput effectiveCalculation() {
    return calculation == null ? new CalculationPolicyInput(null, false) : calculation;
  }
}
