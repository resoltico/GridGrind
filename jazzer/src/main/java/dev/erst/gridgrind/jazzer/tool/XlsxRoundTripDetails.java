package dev.erst.gridgrind.jazzer.tool;

import java.util.Map;

/** Describes a replayed `.xlsx` round-trip input. */
public record XlsxRoundTripDetails(
    int inputBytes, int commandCount, Map<String, Long> commandKinds, Map<String, Long> styleKinds)
    implements ReplayDetails {}
