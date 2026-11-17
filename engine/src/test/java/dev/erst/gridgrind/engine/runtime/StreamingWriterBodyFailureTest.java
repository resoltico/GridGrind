package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempScope;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.MethodSource;

/** Body exceptions and native cleanup failures have independent diagnostic ownership. */
class StreamingWriterBodyFailureTest {
  @TempDir Path root;
  private final AssertionError bodyError = new AssertionError("synthetic fatal serialization");
  private final AssertionError closeError = new AssertionError("synthetic fatal native close");

  static Stream<Arguments> ordinaryFailures() {
    return Stream.of(
        Arguments.of(
            BodyFailure.HOST_ASSERTION, CloseFailure.IO, GridGrindProblemCode.ASSERTION_FAILED),
        Arguments.of(
            BodyFailure.HOST_ASSERTION,
            CloseFailure.UNCHECKED,
            GridGrindProblemCode.ASSERTION_FAILED),
        Arguments.of(BodyFailure.SERIALIZATION, CloseFailure.IO, GridGrindProblemCode.IO_ERROR),
        Arguments.of(
            BodyFailure.SERIALIZATION, CloseFailure.UNCHECKED, GridGrindProblemCode.IO_ERROR));
  }

  @ParameterizedTest
  @MethodSource("ordinaryFailures")
  void retainsEveryBodyFailureAndAppendsNativeCloseEvidence(
      BodyFailure body, CloseFailure close, GridGrindProblemCode expected) throws Exception {
    var normal = run(body, CloseFailure.NONE);
    var failedClose = run(body, close);
    assertEquals(expected, normal.problem().code());
    assertEquals(normal.problem().code(), failedClose.problem().code());
    assertEquals(normal.problem().category(), failedClose.problem().category());
    assertEquals(normal.problem().recovery(), failedClose.problem().recovery());
    assertEquals(normal.problem().title(), failedClose.problem().title());
    assertEquals(normal.problem().message(), failedClose.problem().message());
    assertEquals(normal.problem().resolution(), failedClose.problem().resolution());
    assertEquals(normal.problem().context(), failedClose.problem().context());
    assertEquals(normal.problem().causes().size() + 1, failedClose.problem().causes().size());
    var secondary = failedClose.problem().causes().getLast();
    assertEquals(
        close == CloseFailure.IO
            ? GridGrindProblemCode.IO_ERROR
            : GridGrindProblemCode.INTERNAL_ERROR,
        secondary.code());
    assertEquals("EXECUTE_REQUEST", secondary.stage());
    if (close == CloseFailure.IO)
      assertTrue(secondary.message().contains("synthetic native close"));
    assertEquals(normal.assertions(), failedClose.assertions());
    assertEquals(normal.inspections(), failedClose.inspections());
    assertEquals(normal.calculation(), failedClose.calculation());
    assertEquals(normal.persistence(), failedClose.persistence());
    assertEquals(normal.evidence(), failedClose.evidence());
    var normalJournal =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, normal.journal().outcome());
    var failedCloseJournal =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, failedClose.journal().outcome());
    assertEquals(1, normalJournal.completedStepCount());
    assertEquals(normalJournal.completedStepCount(), failedCloseJournal.completedStepCount());
    assertEquals(normalJournal.failedStep(), failedCloseJournal.failedStep());
    if (body == BodyFailure.HOST_ASSERTION) {
      var evidence =
          assertInstanceOf(
              WorkbookExecutionEvidence.TaskSpecific.Failed.class,
              normal.evidence().taskSpecific());
      assertEquals(1, evidence.assertions().size());
      assertEquals(
          WorkbookExecutionEvidence.TaskSpecific.Origin.HOST_REQUIRED,
          evidence.assertions().getFirst().origin());
    } else {
      assertTrue(normal.problem().message().contains("synthetic serialization unavailable"));
    }
  }

  @ParameterizedTest
  @EnumSource(
      value = CloseFailure.class,
      names = {"IO", "UNCHECKED", "FATAL"})
  void fatalBodyPropagatesWithOneSuppressedNativeCloseFailure(CloseFailure close) {
    var propagated = assertThrows(AssertionError.class, () -> run(BodyFailure.FATAL, close));
    assertSame(bodyError, propagated);
    assertEquals(1, propagated.getSuppressed().length);
    if (close == CloseFailure.FATAL) assertSame(closeError, propagated.getSuppressed()[0]);
    else assertTrue(propagated.getSuppressed()[0].getMessage().contains("synthetic native close"));
  }

  @Test
  void fatalClosePropagatesAndRetainsAnOrdinaryBodyFailureAsSuppressed() {
    var propagated =
        assertThrows(
            AssertionError.class, () -> run(BodyFailure.SERIALIZATION, CloseFailure.FATAL));
    assertSame(closeError, propagated);
    assertEquals(1, propagated.getSuppressed().length);
    assertInstanceOf(IOException.class, propagated.getSuppressed()[0]);
    assertTrue(
        propagated.getSuppressed()[0].getMessage().contains("synthetic serialization unavailable"));
  }

  private WorkbookResult.Failure run(BodyFailure body, CloseFailure close) throws IOException {
    var request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                root.resolve("never-published.xlsx").toString(),
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicySpec.mode(new ExecutionModeSpec.StreamingWrite()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure"),
                    new SheetSelector.ByName(new SheetName("Data")),
                    new WorkbookMutationAction.EnsureSheet())));
    AcceptancePolicy policy =
        body == BodyFailure.HOST_ASSERTION
            ? new AcceptancePolicy.RequireAll(
                List.of(
                    new AcceptancePolicy.Requirement.TerminalAssertions(
                        List.of(
                            new AssertionStep(
                                new StepId("host-required"),
                                new CellSelector.ByAddress(
                                    new SheetName("Data"), new CellAddress("A1")),
                                new CellAssertion.Value(new CellValue.Text("required")))))))
            : AcceptancePolicy.minimum();
    var bindings = ExecutionInputBindingsFixtureSupport.bindings(root, request, policy);
    AtomicInteger closes = new AtomicInteger();
    AtomicInteger publications = new AtomicInteger();
    AtomicInteger serializationAttempts = new AtomicInteger();
    TempFileFactory temp = bindings.tempFileFactory();
    WorkbookTempFileFactory nativeTemp =
        (prefix, suffix) -> {
          if ("poi-sxssf-template".equals(prefix) && body != BodyFailure.HOST_ASSERTION) {
            serializationAttempts.incrementAndGet();
            if (body == BodyFailure.FATAL) throw bodyError;
            throw new IOException("synthetic serialization unavailable");
          }
          return temp.createTempFile(prefix, suffix);
        };
    var engine = new WorkbookExecutionEngine();
    var selectors = new SemanticSelectorResolver(engine);
    var steps =
        new StepExecution(
            engine, selectors, new AssertionExecutor(engine, selectors), temp, bindings);
    var workflow =
        new ExecutionStreamingWorkflow(
            (staged, persistence, source, inputs, acceptance) -> {
              publications.incrementAndGet();
              throw new AssertionError("Failed execution must not publish");
            },
            new CalculationWorkflow(writer -> {}),
            steps,
            temp,
            writer -> {
              closes.incrementAndGet();
              writer.close();
              switch (close) {
                case NONE -> {}
                case IO -> throw new IOException("synthetic native close");
                case UNCHECKED -> throw new IllegalStateException("synthetic native close");
                case FATAL -> throw closeError;
              }
            });
    try (var prepared = RequestPreflight.verify(request, bindings).requirePreparedRequest()) {
      return assertInstanceOf(
          WorkbookResult.Failure.class,
          WorkbookTempScope.call(
              nativeTemp,
              () ->
                  workflow.execute(
                      new ExecutionStreamingWorkflow.Call(
                          ProtocolVersion.CURRENT,
                          prepared,
                          new ExecutionModeSpec.StreamingWrite(),
                          List.of(),
                          ExecutionJournalRecorder.start(
                              request, ExecutionProgressSink.NOOP, root)))));
    } finally {
      bindings.closeRunResources();
      assertEquals(1, closes.get());
      assertEquals(0, publications.get());
      if (body != BodyFailure.HOST_ASSERTION) assertEquals(1, serializationAttempts.get());
    }
  }

  /** Stage that supplies the primary native-work failure. */
  private enum BodyFailure {
    HOST_ASSERTION,
    SERIALIZATION,
    FATAL
  }

  /** Independent native cleanup outcome injected after the real writer closes. */
  private enum CloseFailure {
    NONE,
    IO,
    UNCHECKED,
    FATAL
  }
}
