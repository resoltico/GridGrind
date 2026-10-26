package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.source.*;
import dev.erst.gridgrind.core.vocabulary.*;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableDefinition;
import java.io.*;
import java.nio.file.*;
import java.util.*;

/** Generates table, named-range, and pivot values for operation-sequence fuzzing. */
final class OperationSequenceTabularValues {
  private OperationSequenceTabularValues() {}

  static TableSpec nextTableInput(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return TableSpec.withDefaultMetadata(
        new TableName(tableName),
        new SheetName(sheetName),
        new dev.erst.gridgrind.core.vocabulary.RangeRef(
            validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        data.consumeBoolean(),
        nextTableStyleInput(data));
  }

  static ExcelTableDefinition nextExcelTableDefinition(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return new ExcelTableDefinition(
        tableName,
        sheetName,
        validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false),
        data.consumeBoolean(),
        nextExcelTableStyle(data));
  }

  static TableStyleSpec nextTableStyleInput(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new TableStyleSpec.None();
      default ->
          new TableStyleSpec.Named(
              data.consumeBoolean() ? "TableStyleMedium2" : "TableStyleLight9",
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean());
    };
  }

  static ExcelTableStyle nextExcelTableStyle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelTableStyle.None();
      default ->
          new ExcelTableStyle.Named(
              data.consumeBoolean() ? "TableStyleMedium2" : "TableStyleLight9",
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean());
    };
  }

  static TableSelector nextTableSelector(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new TableSelector.All();
      default ->
          new TableSelector.ByNames(
              List.of(
                  new TableName(
                      data.consumeBoolean()
                          ? nextTableName(data, true, primarySheet)
                          : nextTableName(data, true, secondarySheet))));
    };
  }

  static PivotTableSelector nextPivotTableSelector(
      GridGrindFuzzData data,
      String primarySheet,
      String secondarySheet,
      String pivotTableName,
      boolean validName) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new PivotTableSelector.All();
      default ->
          data.consumeBoolean()
              ? new PivotTableSelector.ByNameOnSheet(
                  validName ? pivotTableName : nextPivotTableName(data, false),
                  new SheetName(primarySheet))
              : new PivotTableSelector.ByNameOnSheet(
                  validName ? pivotTableName : nextPivotTableName(data, false),
                  new SheetName(secondarySheet));
    };
  }

  static PivotTableSpec nextPivotTableInput(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return new PivotTableSpec(
        validName ? pivotTableName : nextPivotTableName(data, false),
        new SheetName(targetSheet),
        nextPivotTableSource(data, targetSheet, namedRangeName, tableName, validName, validRange),
        new PivotTableSpec.Anchor(
            new dev.erst.gridgrind.core.vocabulary.CellAddress(
                data.consumeBoolean() ? "F4" : "A6")),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new PivotTableSpec.DataField(
                "Actual",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Actual",
                Optional.empty())));
  }

  static PivotTableSpec.PivotSourceSpec nextPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new PivotTableSpec.PivotSourceSpec.Range(
              new SheetName(targetSheet),
              new dev.erst.gridgrind.core.vocabulary.RangeRef(
                  validRange ? "A1:C4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
      case 1 ->
          new PivotTableSpec.PivotSourceSpec.NamedRange(
              new dev.erst.gridgrind.core.vocabulary.DefinedName(
                  validName ? namedRangeName : nextNamedRangeName(data, false)));
      default ->
          new PivotTableSpec.PivotSourceSpec.Table(
              new dev.erst.gridgrind.core.vocabulary.TableName(
                  validName ? tableName : nextTableName(data, false, targetSheet)));
    };
  }

  static ExcelPivotTableDefinition nextExcelPivotTableDefinition(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return new ExcelPivotTableDefinition(
        validName ? pivotTableName : nextPivotTableName(data, false),
        targetSheet,
        nextExcelPivotTableSource(
            data, targetSheet, namedRangeName, tableName, validName, validRange),
        new ExcelPivotTableDefinition.Anchor(data.consumeBoolean() ? "F4" : "A6"),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new ExcelPivotTableDefinition.DataField(
                "Actual",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Actual",
                Optional.empty())));
  }

  static ExcelPivotTableDefinition.Source nextExcelPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new ExcelPivotTableDefinition.Source.Range(
              targetSheet, validRange ? "A1:C4" : FuzzDataDecoders.nextNonBlankRange(data, false));
      case 1 ->
          new ExcelPivotTableDefinition.Source.NamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false));
      default ->
          new ExcelPivotTableDefinition.Source.Table(
              validName ? tableName : nextTableName(data, false, targetSheet));
    };
  }

  static String nextTableName(GridGrindFuzzData data, boolean valid, String sheetName) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (!valid) {
      return nextNamedRangeName(data, false);
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> sheetName + "Table";
      case 1 -> "BudgetTable";
      default -> "OpsTable";
    };
  }

  static String nextNamedRangeName(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");
    if (!valid) {
      return switch (selectorSlot(nextSelectorByte(data))) {
        case 0 -> "";
        case 1 -> "A1";
        case 2 -> "R1C1";
        case 3 -> "_xlnm.Print_Area";
        default -> "1Budget";
      };
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> "BudgetTotal";
      case 1 -> "LocalItem";
      case 2 -> "Report_Value";
      case 3 -> "Summary.Total";
      default -> "Name" + data.consumeInt(1, 9);
    };
  }

  static String nextPivotTableName(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");
    if (!valid) {
      return "";
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> PIVOT_TABLE_NAME;
      case 1 -> "Budget Pivot";
      default -> "Pivot " + data.consumeInt(1, 9);
    };
  }
}
