package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Objects;

/** Factory for the engine module's narrow published request-execution seam. */
public final class GridGrindEngine {
  private static final GridGrindRequestExecutor REQUEST_EXECUTOR =
      new ProductionRequestExecutor(
          new dev.erst.gridgrind.engine.runtime.DefaultGridGrindRequestExecutor());
  private static final GridGrindRequestDoctor REQUEST_DOCTOR =
      new ProductionRequestDoctor(new dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor());

  private GridGrindEngine() {}

  /** Returns the production request executor backed by the workbook engine runtime. */
  public static GridGrindRequestExecutor requestExecutor() {
    return REQUEST_EXECUTOR;
  }

  /** Returns the production request doctor backed by the workbook engine runtime. */
  public static GridGrindRequestDoctor requestDoctor() {
    return REQUEST_DOCTOR;
  }

  private record ProductionRequestExecutor(
      dev.erst.gridgrind.engine.runtime.DefaultGridGrindRequestExecutor delegate)
      implements GridGrindRequestExecutor {
    private ProductionRequestExecutor {
      Objects.requireNonNull(delegate, "delegate must not be null");
    }

    @Override
    public dev.erst.gridgrind.contract.dto.GridGrindResponse execute(
        WorkbookPlan request, GridGrindRequestInputs inputs, GridGrindJournalSink sink) {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(inputs, "inputs must not be null");
      Objects.requireNonNull(sink, "sink must not be null");
      return delegate.execute(
          request,
          toInternalInputs(inputs),
          event -> GridGrindJournalSink.requireNonNull(sink).emit(event));
    }
  }

  private record ProductionRequestDoctor(
      dev.erst.gridgrind.engine.runtime.GridGrindRequestDoctor delegate)
      implements GridGrindRequestDoctor {
    private ProductionRequestDoctor {
      Objects.requireNonNull(delegate, "delegate must not be null");
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request) {
      return delegate.diagnose(request);
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs) {
      Objects.requireNonNull(inputs, "inputs must not be null");
      return delegate.diagnose(request, toInternalInputs(inputs));
    }
  }

  private static dev.erst.gridgrind.engine.runtime.ExecutionInputBindings toInternalInputs(
      GridGrindRequestInputs inputs) {
    Objects.requireNonNull(inputs, "inputs must not be null");
    return inputs
        .standardInputBytes()
        .map(
            bytes ->
                new dev.erst.gridgrind.engine.runtime.ExecutionInputBindings(
                    inputs.workingDirectory(), inputs.tempRoot(), bytes))
        .orElseGet(
            () ->
                new dev.erst.gridgrind.engine.runtime.ExecutionInputBindings(
                    inputs.workingDirectory(), inputs.tempRoot()));
  }
}
