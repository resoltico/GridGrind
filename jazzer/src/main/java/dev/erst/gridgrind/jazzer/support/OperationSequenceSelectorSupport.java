package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.formula.FormulaCellTarget;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;

/** Shared selector, span, and execution-policy helpers for Jazzer operation generation. */
final class OperationSequenceSelectorSupport {
  private OperationSequenceSelectorSupport() {}

  static boolean validFormulaAddress(GridGrindFuzzData data) {
    return data.consumeBoolean();
  }

  static ExecutionPolicyInput nextExecutionPolicy(
      GridGrindFuzzData data, String primarySheet, String secondarySheet, boolean validAddress) {
    return switch (selectorSlot(nextSelectorByte(data)) % 6) {
      case 0 -> ExecutionPolicyInput.defaults();
      case 1 -> ProtocolStepSupport.executionPolicy(ProtocolStepSupport.calculateAll());
      case 2 ->
          ProtocolStepSupport.executionPolicy(
              ProtocolStepSupport.calculateTargets(
                  nextFormulaTarget(data, primarySheet, secondarySheet, validAddress)));
      case 3 -> ProtocolStepSupport.executionPolicy(ProtocolStepSupport.clearFormulaCaches());
      case 4 -> ProtocolStepSupport.executionPolicy(ProtocolStepSupport.markRecalculateOnOpen());
      default ->
          ProtocolStepSupport.executionPolicy(
              ProtocolStepSupport.calculateAllAndMarkRecalculateOnOpen());
    };
  }

  static FormulaCellTarget nextFormulaTarget(
      GridGrindFuzzData data, String primarySheet, String secondarySheet, boolean validAddress) {
    return new FormulaCellTarget(
        new dev.erst.gridgrind.core.vocabulary.SheetName(
            data.consumeBoolean() ? primarySheet : secondarySheet),
        new CellAddress(nextFormulaTargetAddress(data, validAddress)));
  }

  static String nextFormulaTargetAddress(GridGrindFuzzData data, boolean validAddress) {
    return validAddress ? "C2" : FuzzDataDecoders.nextNonBlankCellAddress(data, false);
  }

  static RangeSelector nextConditionalFormattingTarget(
      GridGrindFuzzData data, String sheetName, boolean validRange) {
    String firstRange = validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false);
    if (data.consumeBoolean()) {
      return new RangeSelector.ByRange(new SheetName(sheetName), new RangeRef(firstRange));
    }
    String secondRange = validRange ? "C1:C4" : FuzzDataDecoders.nextNonBlankRange(data, false);
    return new RangeSelector.ByRanges(
        new SheetName(sheetName), List.of(new RangeRef(firstRange), new RangeRef(secondRange)));
  }

  static IndexSpan nextIndexSpan(GridGrindFuzzData data, int upperBound) {
    int first = data.consumeInt(0, upperBound - 1);
    return new IndexSpan(first, data.consumeInt(first, upperBound));
  }

  static RowBandSelector.Span nextRowHeightTarget(
      GridGrindFuzzData data, String sheetName, IndexSpan boundedSpan) {
    return data.consumeBoolean()
        ? new RowBandSelector.Span(
            new SheetName(sheetName), boundedSpan.first(), boundedSpan.last())
        : new RowBandSelector.Span(new SheetName(sheetName), 0, 250_000);
  }

  static int nextNonZeroDelta(GridGrindFuzzData data, int upperBound) {
    int absoluteDelta = data.consumeInt(1, upperBound);
    return data.consumeBoolean() ? absoluteDelta : -absoluteDelta;
  }

  static int nextSelectorByte(GridGrindFuzzData data) {
    return Byte.toUnsignedInt(data.consumeByte());
  }

  static int selectorFamily(int selector) {
    return selector >>> 4;
  }

  static int selectorSlot(int selector) {
    return selector & 0x0F;
  }

  record IndexSpan(int first, int last) {}
}
