package dev.erst.gridgrind.jazzer.tool;

import java.util.Map;

/** Describes a replayed raw protocol-request fuzz input. */
public record ProtocolRequestDetails(
    int inputBytes,
    String decodeOutcome,
    String sourceKind,
    String persistenceKind,
    int operationCount,
    Map<String, Long> operationKinds,
    Map<String, Long> styleKinds,
    int analysisSheetCount)
    implements ReplayDetails {}
