package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.engine.bootstrap.NativeProviderInvocation;

/** Factory for the engine's published ports; each call owns a fresh synchronous native realm. */
public final class GridGrindEngine {
  private static final GridGrindRequestExecutor REQUEST_EXECUTOR = new ProductionRequestExecutor();
  private static final GridGrindRequestDoctor REQUEST_DOCTOR = new ProductionRequestDoctor();

  private GridGrindEngine() {}

  /** Returns the reusable execution port; each invocation owns its native realm. */
  public static GridGrindRequestExecutor requestExecutor() {
    return REQUEST_EXECUTOR;
  }

  /** Returns the reusable doctor port; each diagnosis owns its native realm. */
  public static GridGrindRequestDoctor requestDoctor() {
    return REQUEST_DOCTOR;
  }

  /** Creates an invocation-owned implementation for each execution call. */
  private static final class ProductionRequestExecutor implements GridGrindRequestExecutor {
    @Override
    public WorkbookResult execute(
        WorkbookPlan request, GridGrindRequestInputs inputs, GridGrindProgressSink sink) {
      return NativeProviderInvocation.create(GridGrindRequestExecutor.class)
          .execute(request, inputs, sink);
    }

    @Override
    public WorkbookResult execute(
        RequestIntake analysis, GridGrindRequestInputs inputs, GridGrindProgressSink sink) {
      return NativeProviderInvocation.create(GridGrindRequestExecutor.class)
          .execute(analysis, inputs, sink);
    }
  }

  /** Creates an invocation-owned implementation for each diagnosis call. */
  private static final class ProductionRequestDoctor implements GridGrindRequestDoctor {
    @Override
    public RequestDoctorReport diagnose(RequestIntake analysis, RequestInput input) {
      return NativeProviderInvocation.create(GridGrindRequestDoctor.class)
          .diagnose(analysis, input);
    }

    @Override
    public RequestDoctorReport diagnose(
        RequestIntake analysis, RequestInput input, GridGrindRequestInputs inputs) {
      return NativeProviderInvocation.create(GridGrindRequestDoctor.class)
          .diagnose(analysis, input, inputs);
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request) {
      return NativeProviderInvocation.create(GridGrindRequestDoctor.class).diagnose(request);
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs) {
      return NativeProviderInvocation.create(GridGrindRequestDoctor.class)
          .diagnose(request, inputs);
    }
  }
}
