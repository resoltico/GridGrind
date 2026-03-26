package dev.erst.gridgrind.jazzer.tool;

import java.util.Map;

/** Describes the structured meaning of a replayed fuzz input. */
public sealed interface ReplayDetails
    permits CommandSequenceDetails,
        ProtocolRequestDetails,
        ProtocolWorkflowDetails,
        XlsxRoundTripDetails {}
