package dev.erst.gridgrind.contract.assertion;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.catalog.ProtocolTargetingMode;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.Objects;
import java.util.Optional;

/** Assertions that derive their target selectors from nested analysis queries. */
public sealed interface AnalysisAssertion extends Assertion
    permits AnalysisAssertion.AnalysisMaxSeverity,
        AnalysisAssertion.AnalysisFindingPresent,
        AnalysisAssertion.AnalysisFindingAbsent {

  String ANALYSIS_RULE = "Matches the nested analysis query's target selectors.";

  @ProtocolTypeMetadata(
      id = "EXPECT_ANALYSIS_MAX_SEVERITY",
      summary =
          "Run one analysis query and require its highest severity to be no higher than maximumSeverity.",
      targetingMode = ProtocolTargetingMode.ANALYSIS_QUERY,
      targetSelectorRule = ANALYSIS_RULE)
  record AnalysisMaxSeverity(InspectionQuery query, AnalysisSeverity maximumSeverity)
      implements AnalysisAssertion {
    public AnalysisMaxSeverity {
      query = AssertionSupport.requireAnalysisQuery(query, "query");
      maximumSeverity = AssertionSupport.requireSeverity(maximumSeverity, "maximumSeverity");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_ANALYSIS_FINDING_PRESENT",
      summary = "Run one analysis query and require at least one matching finding.",
      targetingMode = ProtocolTargetingMode.ANALYSIS_QUERY,
      targetSelectorRule = ANALYSIS_RULE)
  record AnalysisFindingPresent(
      InspectionQuery query,
      AnalysisFindingCode code,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<AnalysisSeverity> severity,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> messageContains)
      implements AnalysisAssertion {
    public AnalysisFindingPresent {
      query = AssertionSupport.requireAnalysisQuery(query, "query");
      code = AssertionSupport.requireFindingCode(code, "code");
      Objects.requireNonNull(severity, "severity must not be null");
      Objects.requireNonNull(messageContains, "messageContains must not be null");
      messageContains =
          messageContains.map(
              value -> {
                if (value.isBlank()) {
                  throw new IllegalArgumentException("messageContains must not be blank");
                }
                return value;
              });
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_ANALYSIS_FINDING_ABSENT",
      summary = "Run one analysis query and require no matching finding.",
      targetingMode = ProtocolTargetingMode.ANALYSIS_QUERY,
      targetSelectorRule = ANALYSIS_RULE)
  record AnalysisFindingAbsent(
      InspectionQuery query,
      AnalysisFindingCode code,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<AnalysisSeverity> severity,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> messageContains)
      implements AnalysisAssertion {
    public AnalysisFindingAbsent {
      query = AssertionSupport.requireAnalysisQuery(query, "query");
      code = AssertionSupport.requireFindingCode(code, "code");
      Objects.requireNonNull(severity, "severity must not be null");
      Objects.requireNonNull(messageContains, "messageContains must not be null");
      messageContains =
          messageContains.map(
              value -> {
                if (value.isBlank()) {
                  throw new IllegalArgumentException("messageContains must not be blank");
                }
                return value;
              });
    }
  }

  /** Returns the selector types accepted by one analysis assertion instance. */
  public static Class<? extends Selector>[] targetSelectorsFor(AnalysisAssertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return switch (assertion) {
      case AnalysisMaxSeverity analysisMaxSeverity ->
          InspectionQuery.allowedTargetTypes(analysisMaxSeverity.query());
      case AnalysisFindingPresent analysisFindingPresent ->
          InspectionQuery.allowedTargetTypes(analysisFindingPresent.query());
      case AnalysisFindingAbsent analysisFindingAbsent ->
          InspectionQuery.allowedTargetTypes(analysisFindingAbsent.query());
    };
  }
}
