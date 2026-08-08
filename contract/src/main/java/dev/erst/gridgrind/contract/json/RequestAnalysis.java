package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Complete structural analysis of one request document before static or dynamic validation. */
public final class RequestAnalysis {
  private final RequestBoundFragments boundFragments;
  private final List<RequestStructuralProblem> structuralProblems;
  private final List<RequestBindingFailure> bindingFailures;
  private final Optional<WorkbookPlan> completePlan;
  private final RequestDiagnosticRedactor diagnosticRedactor;

  /**
   * Creates a manually assembled analysis for focused callers that do not retain a raw parse tree.
   *
   * <p>Such analyses can expose fragments and structural problems but cannot locate a final
   * cross-fragment constructor failure within an original request document.
   */
  public RequestAnalysis(
      RequestBoundFragments boundFragments, List<RequestStructuralProblem> structuralProblems) {
    this(
        boundFragments,
        structuralProblems,
        List.of(),
        RequestDiagnosticRedactor.empty(),
        Optional.empty());
  }

  RequestAnalysis(
      RequestBoundFragments boundFragments,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures,
      RequestDiagnosticRedactor diagnosticRedactor,
      Optional<RequestJsonNode> rawRoot) {
    this.boundFragments = Objects.requireNonNull(boundFragments, "boundFragments must not be null");
    this.structuralProblems =
        RequestStructuralProblemOrder.order(
            List.copyOf(
                Objects.requireNonNull(structuralProblems, "structuralProblems must not be null")));
    this.diagnosticRedactor =
        Objects.requireNonNull(diagnosticRedactor, "diagnosticRedactor must not be null");
    Optional<RequestJsonNode> parsedRoot = Objects.requireNonNullElseGet(rawRoot, Optional::empty);
    List<RequestBindingFailure> discoveredBindingFailures =
        new java.util.ArrayList<>(
            Objects.requireNonNull(bindingFailures, "bindingFailures must not be null"));
    this.completePlan = bindCompletePlan(discoveredBindingFailures);
    this.bindingFailures =
        discoveredBindingFailures.stream()
            .map(failure -> locateBindingFailure(failure, parsedRoot))
            .toList();
  }

  /** Returns each independently bound request fragment, including any valid sibling subtrees. */
  public RequestBoundFragments boundFragments() {
    return boundFragments;
  }

  /** Returns every ordered structural problem discovered while parsing the request once. */
  public List<RequestStructuralProblem> structuralProblems() {
    return structuralProblems;
  }

  /** Returns every independent constructor-level failure discovered after structural collection. */
  public List<RequestBindingFailure> bindingFailures() {
    return bindingFailures;
  }

  /** Returns the request-scoped redactor derived from this analysis's one tolerant parse. */
  public RequestDiagnosticRedactor diagnosticRedactor() {
    return diagnosticRedactor;
  }

  /** Returns the complete typed plan only when every structural fragment and constructor binds. */
  public Optional<WorkbookPlan> completePlan() {
    return completePlan;
  }

  /** Returns the complete plan after a caller has handled every intake problem in this analysis. */
  public WorkbookPlan requireCompletePlan() {
    if (!structuralProblems.isEmpty()) {
      throw new IllegalStateException(
          "A structurally invalid request cannot produce a complete plan");
    }
    if (!bindingFailures.isEmpty()) {
      throw bindingFailures.getFirst().exception();
    }
    return completePlan.orElseThrow(
        () ->
            new IllegalStateException(
                "A structurally valid request must bind every required request fragment"));
  }

  /** Returns whether tolerant parsing and structural collection found no request-shape defects. */
  public boolean isStructurallyValid() {
    return structuralProblems.isEmpty();
  }

  /** Returns whether every structurally valid fragment and the complete plan bound successfully. */
  public boolean isBindable() {
    return structuralProblems.isEmpty() && bindingFailures.isEmpty() && completePlan.isPresent();
  }

  private Optional<WorkbookPlan> bindCompletePlan(
      List<RequestBindingFailure> discoveredBindingFailures) {
    if (!structuralProblems.isEmpty() || !discoveredBindingFailures.isEmpty()) {
      return Optional.empty();
    }
    try {
      return boundFragments.completePlan();
    } catch (IllegalArgumentException exception) {
      // Cross-fragment rules such as duplicate step identifiers have no independently bindable
      // owner, so they are collected after every fragment has successfully bound.
      discoveredBindingFailures.add(
          RequestBindingFailure.fromCompletePlan(exception, diagnosticRedactor));
      return Optional.empty();
    }
  }

  private static RequestBindingFailure locateBindingFailure(
      RequestBindingFailure failure, Optional<RequestJsonNode> rawRoot) {
    Optional<Long> byteOffset =
        rawRoot.flatMap(
            root -> RequestJsonTokenLocationSupport.byteOffsetAt(root, failure.jsonPath()));
    return byteOffset.isPresent() ? failure.locatedAt(byteOffset) : failure;
  }
}
