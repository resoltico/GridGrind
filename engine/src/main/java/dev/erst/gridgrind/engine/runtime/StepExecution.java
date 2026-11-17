package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.PreparedAssertionStep;
import dev.erst.gridgrind.core.plan.PreparedInspectionStep;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import dev.erst.gridgrind.engine.backend.eventread.EventWorkbookReader;
import dev.erst.gridgrind.engine.backend.streaming.StreamingWorkbookWriter;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Executes workbook steps and materialized-read seams for the request executor. */
final class StepExecution {
  private final WorkbookExecutionEngine workbookEngine;
  private final SemanticSelectorResolver selectorResolver;
  private final AssertionExecutor assertionExecutor;
  private final TempFileFactory tempFileFactory;
  private final FormulaAwareMutationStepExecutor mutationStepExecutor;
  private final ExecutionInputBindings bindings;
  private final MaterializedWorkbookOpener materializedOpener;

  StepExecution(
      WorkbookExecutionEngine workbookEngine,
      SemanticSelectorResolver selectorResolver,
      AssertionExecutor assertionExecutor,
      TempFileFactory tempFileFactory,
      ExecutionInputBindings bindings) {
    this(
        workbookEngine,
        selectorResolver,
        assertionExecutor,
        tempFileFactory,
        bindings,
        path ->
            XssfWorkbooks.open(
                path,
                dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment.defaults(),
                tempFileFactory::createTempFile));
  }

  StepExecution(
      WorkbookExecutionEngine workbookEngine,
      SemanticSelectorResolver selectorResolver,
      AssertionExecutor assertionExecutor,
      TempFileFactory tempFileFactory,
      ExecutionInputBindings bindings,
      MaterializedWorkbookOpener materializedOpener) {
    this.bindings = Objects.requireNonNull(bindings);
    this.materializedOpener = Objects.requireNonNull(materializedOpener);
    this.workbookEngine = Objects.requireNonNull(workbookEngine, "workbookEngine must not be null");
    this.selectorResolver =
        Objects.requireNonNull(selectorResolver, "selectorResolver must not be null");
    this.assertionExecutor =
        Objects.requireNonNull(assertionExecutor, "assertionExecutor must not be null");
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    this.mutationStepExecutor =
        new FormulaAwareMutationStepExecutor(this.selectorResolver, bindings);
  }

  FormulaAwareMutationStepExecutor mutationStepExecutor() {
    return mutationStepExecutor;
  }

  InspectionResult executeInspectionStep(
      PreparedInspectionStep inspectionStep,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeSpec.EventRead)) {
      return executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
    }
    return executeEventInspectionAgainstWorkbook(inspectionStep, workbook);
  }

  AssertionResult executeAssertionStep(
      PreparedAssertionStep assertionStep,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode)
      throws IOException, AssertionFailedException {
    if (!(executionMode instanceof ExecutionModeSpec.EventRead)) {
      return assertionExecutor.execute(assertionStep, workbook, workbookLocation);
    }
    throw new IllegalStateException(
        "execution.mode.type=EVENT_READ does not support assertion steps");
  }

  AssertionStepExecution executeAssertionStepCollecting(
      PreparedAssertionStep assertionStep,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeSpec.EventRead)) {
      return assertionExecutor.executeCollecting(assertionStep, workbook, workbookLocation);
    }
    throw new IllegalStateException(
        "execution.mode.type=EVENT_READ does not support assertion steps");
  }

  private InspectionResult executeFullInspectionStep(
      PreparedInspectionStep inspectionStep,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SemanticSelectorResolver.ResolvedInspectionTarget resolvedTarget =
        selectorResolver.resolveInspectionTarget(
            inspectionStep.authoredStep().stepId(),
            workbook,
            inspectionStep.target(),
            inspectionStep.authoredStep().query());
    if (resolvedTarget.isShortCircuit()) {
      return Objects.requireNonNull(
          resolvedTarget.shortCircuitResult(), "short-circuit inspection result must be present");
    }
    dev.erst.gridgrind.core.plan.InspectionResult result =
        workbookEngine
            .read(
                workbook,
                workbookLocation,
                resolvedInspection(
                    inspectionStep.authoredStep().stepId(),
                    Objects.requireNonNull(
                        resolvedTarget.selector(), "resolved selector must be present"),
                    inspectionStep.authoredStep().query()))
            .getFirst();
    return result;
  }

  private InspectionResult executeEventInspectionAgainstWorkbook(
      PreparedInspectionStep inspectionStep, XssfWorkbook workbook) throws IOException {
    Path tempPath = null;
    try {
      tempPath =
          WorkbookTempFileFactory.prepareCreateNewTarget(
              tempFileFactory.createTempFile("gridgrind-event-read-", ".xlsx"));
      workbook
          .persistence()
          .savePlainWorkbook(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeEventInspection(tempPath, inspectionStep);
    } finally {
      WorkbookLifecycle.deleteIfExists(tempPath);
    }
  }

  MaterializedReadView openReadView(Path path, WorkbookLocation location) throws IOException {
    return new MaterializedReadView(materializedOpener.open(path), location);
  }

  StreamingReadView streamingReadView(StreamingWorkbookWriter writer, WorkbookLocation location) {
    var view = new StreamingReadView(writer, location);
    bindings.own(view::close);
    return view;
  }

  /** One XSSF view shared only by observations of the same authored workbook state. */
  final class MaterializedReadView implements AutoCloseable {
    private final XssfWorkbook workbook;
    private final WorkbookLocation location;

    private MaterializedReadView(XssfWorkbook workbook, WorkbookLocation location) {
      this.workbook = workbook;
      this.location = location;
    }

    InspectionResult inspect(PreparedInspectionStep step) {
      return executeFullInspectionStep(step, workbook, location);
    }

    AssertionResult assertThat(PreparedAssertionStep step) throws AssertionFailedException {
      return assertionExecutor.execute(step, workbook, location);
    }

    AssertionStepExecution collect(PreparedAssertionStep step) {
      return assertionExecutor.executeCollecting(step, workbook, location);
    }

    @Override
    public void close() throws IOException {
      workbook.close();
    }
  }

  /** Lazily snapshots the completed streaming state for its terminal observations. */
  final class StreamingReadView implements AutoCloseable {
    private final StreamingWorkbookWriter writer;
    private final WorkbookLocation location;
    private @Nullable Path path;
    private @Nullable MaterializedReadView view;

    private StreamingReadView(StreamingWorkbookWriter writer, WorkbookLocation location) {
      this.writer = writer;
      this.location = location;
    }

    InspectionResult inspect(PreparedInspectionStep step) throws IOException {
      return view().inspect(step);
    }

    AssertionResult assertThat(PreparedAssertionStep step)
        throws IOException, AssertionFailedException {
      return view().assertThat(step);
    }

    AssertionStepExecution collect(PreparedAssertionStep step) throws IOException {
      return view().collect(step);
    }

    private MaterializedReadView view() throws IOException {
      if (view == null) {
        Path created =
            WorkbookTempFileFactory.prepareCreateNewTarget(
                tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx"));
        path = created;
        writer.save(created, WorkbookArtifactWriteDisposition.CREATE_NEW);
        view = openReadView(created, location);
      }
      return view;
    }

    @SuppressWarnings(
        "PMD.CloseResource") // The drained nullable handle is explicitly closed below on every
    // nonnull path.
    @Override
    public void close() throws IOException {
      MaterializedReadView opened = view;
      view = null;
      Path owned = path;
      path = null;
      IOException failure = null;
      if (opened != null) {
        try {
          opened.close();
        } catch (IOException exception) {
          failure = exception;
        }
      }
      if (owned != null) {
        try {
          java.nio.file.Files.deleteIfExists(owned);
        } catch (IOException exception) {
          if (failure == null) failure = exception;
          else failure.addSuppressed(exception);
        }
      }
      if (failure != null) throw failure;
    }
  }

  /** Native opening boundary for immutable materialized workbook views. */
  @FunctionalInterface
  interface MaterializedWorkbookOpener {
    /** Opens one native resource whose close obligation transfers to the receiving read view. */
    XssfWorkbook open(Path path) throws IOException;
  }

  InspectionResult executeEventInspection(Path workbookPath, PreparedInspectionStep inspectionStep)
      throws IOException {
    EventWorkbookReader eventWorkbookReader = new EventWorkbookReader();
    dev.erst.gridgrind.core.plan.ResolvedInspectionStep command =
        resolvedInspection(inspectionStep);
    if (!(command.query() instanceof dev.erst.gridgrind.core.plan.InspectionQuery.Introspection)) {
      throw new IllegalArgumentException(
          BackendCapabilities.eventRead()
              .unsupportedQueryMessage(inspectionStep.authoredStep().query().queryType()));
    }
    dev.erst.gridgrind.core.plan.InspectionResult result =
        eventWorkbookReader.apply(workbookPath, List.of(command)).getFirst();
    return result;
  }

  private static dev.erst.gridgrind.core.plan.ResolvedInspectionStep resolvedInspection(
      dev.erst.gridgrind.core.vocabulary.StepId id,
      dev.erst.gridgrind.core.plan.Selector target,
      dev.erst.gridgrind.core.plan.InspectionQuery query) {
    return new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
        id, (dev.erst.gridgrind.core.plan.ResolvedTarget) target, query);
  }

  private static dev.erst.gridgrind.core.plan.ResolvedInspectionStep resolvedInspection(
      dev.erst.gridgrind.core.plan.PreparedInspectionStep step) {
    return resolvedInspection(
        step.authoredStep().stepId(), step.target(), step.authoredStep().query());
  }
}
