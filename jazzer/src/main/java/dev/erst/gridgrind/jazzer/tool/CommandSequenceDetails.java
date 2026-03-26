package dev.erst.gridgrind.jazzer.tool;

import java.util.Map;

/** Describes a replayed workbook-command sequence. */
public record CommandSequenceDetails(
    int inputBytes, int commandCount, Map<String, Long> commandKinds, Map<String, Long> styleKinds)
    implements ReplayDetails {}
