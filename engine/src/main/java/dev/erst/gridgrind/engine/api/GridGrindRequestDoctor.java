package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Objects;

/** Transport-neutral request-doctor surface for one authored workbook plan. */
public interface GridGrindRequestDoctor {
  /** Returns one machine-readable lint report for the supplied request. */
  RequestDoctorReport diagnose(WorkbookPlan request);

  /**
   * Returns one machine-readable lint report for the supplied request using explicit authored
   * inputs when input resolution should be validated as part of linting.
   */
  RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs);

  /** Returns a doctor that rejects null delegates up front. */
  static GridGrindRequestDoctor requireNonNull(GridGrindRequestDoctor doctor) {
    return Objects.requireNonNull(doctor, "doctor must not be null");
  }
}
