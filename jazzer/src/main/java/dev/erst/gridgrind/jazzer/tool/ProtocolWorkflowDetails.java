package dev.erst.gridgrind.jazzer.tool;

import java.util.Map;

/** Describes a replayed structured protocol workflow. */
public record ProtocolWorkflowDetails(
    int inputBytes,
    String sourceKind,
    String persistenceKind,
    int operationCount,
    Map<String, Long> operationKinds,
    Map<String, Long> styleKinds,
    int assertionCount,
    Map<String, Long> assertionKinds,
    int readCount,
    Map<String, Long> readKinds,
    String responseKind)
    implements ReplayDetails {}
