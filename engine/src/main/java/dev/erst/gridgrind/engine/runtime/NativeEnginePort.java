package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.semantics.ExecutionConventions;
import dev.erst.gridgrind.engine.api.GridGrindProgressSink;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.engine.poi.NativeInvocationContext;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Supplier;

/** One owned native invocation, reached exclusively through the shared published ports. */
public final class NativeEnginePort
    implements dev.erst.gridgrind.engine.api.GridGrindRequestExecutor,
        dev.erst.gridgrind.engine.api.GridGrindRequestDoctor {
  private final NativeInvocationContext context;
  private final DefaultGridGrindRequestExecutor executor;
  private final GridGrindRequestDoctor doctor;

  /** Enters the invocation-owned provider policy before constructing native adapters. */
  public NativeEnginePort(ExecutionConventions conventions) {
    context = NativeInvocationContext.enterOwned(conventions);
    try {
      executor = new DefaultGridGrindRequestExecutor(context);
      doctor = new GridGrindRequestDoctor(context);
    } catch (RuntimeException | Error primary) {
      context.closePreserving(primary);
      throw primary;
    }
  }

  private static ExecutionInputBindings inputs(GridGrindRequestInputs inputs) {
    Objects.requireNonNull(inputs);
    return new ExecutionInputBindings(
        inputs.workingDirectory(),
        inputs.tempRoot(),
        inputs.executionGrant(),
        inputs.standardInputSource(),
        inputs.secretResolver());
  }

  @Override
  public WorkbookResult execute(
      WorkbookPlan request, GridGrindRequestInputs inputs, GridGrindProgressSink sink) {
    return executeOwned(() ->
        executor.execute(
            Objects.requireNonNull(request), inputs(inputs), Objects.requireNonNull(sink)::emit));
  }

  @Override
  public WorkbookResult execute(
      RequestIntake analysis, GridGrindRequestInputs inputs, GridGrindProgressSink sink) {
    return executeOwned(() ->
        executor.execute(
            Objects.requireNonNull(analysis).requireCompletePlan(),
            inputs(inputs),
            Objects.requireNonNull(sink)::emit,
            Optional.of(analysis)));
  }

  private WorkbookResult executeOwned(Supplier<WorkbookResult> work) {
    context.requireOpen();
    try {
      return work.get();
    } catch (RuntimeException | Error primary) {
      context.closePreserving(primary);
      throw primary;
    }
  }

  @Override
  public RequestDoctorReport diagnose(RequestIntake analysis, RequestInput input) {
    return diagnoseOwned(() -> doctor.diagnose(analysis, input));
  }

  @Override
  public RequestDoctorReport diagnose(
      RequestIntake analysis, RequestInput input, GridGrindRequestInputs inputs) {
    return diagnoseOwned(() -> doctor.diagnose(analysis, input, inputs(inputs)));
  }

  @Override
  public RequestDoctorReport diagnose(WorkbookPlan request) {
    return diagnoseOwned(() -> doctor.diagnose(request));
  }

  @Override
  public RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs) {
    return diagnoseOwned(() -> doctor.diagnose(request, inputs(inputs)));
  }

  private RequestDoctorReport diagnoseOwned(Supplier<RequestDoctorReport> work) {
    context.requireOpen();
    try {
      return work.get();
    } catch (RuntimeException | Error primary) {
      context.closePreserving(primary);
      throw primary;
    }
  }
}
