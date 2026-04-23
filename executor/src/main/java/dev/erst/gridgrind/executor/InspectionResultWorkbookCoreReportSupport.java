package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSnapshot;

/** Converts workbook-global security, summary, protection, and named-range snapshots. */
final class InspectionResultWorkbookCoreReportSupport {
  private InspectionResultWorkbookCoreReportSupport() {}

  static GridGrindResponse.WorkbookSummary toWorkbookSummary(
      dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary workbookSummary) {
    return switch (workbookSummary) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty empty ->
          new GridGrindResponse.WorkbookSummary.Empty(
              empty.sheetCount(),
              empty.sheetNames(),
              empty.namedRangeCount(),
              empty.forceFormulaRecalculationOnOpen());
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets withSheets ->
          new GridGrindResponse.WorkbookSummary.WithSheets(
              withSheets.sheetCount(),
              withSheets.sheetNames(),
              withSheets.activeSheetName(),
              withSheets.selectedSheetNames(),
              withSheets.namedRangeCount(),
              withSheets.forceFormulaRecalculationOnOpen());
    };
  }

  static OoxmlPackageSecurityReport toOoxmlPackageSecurityReport(
      dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySnapshot snapshot) {
    return new OoxmlPackageSecurityReport(
        toOoxmlEncryptionReport(snapshot.encryption()),
        snapshot.signatures().stream()
            .map(InspectionResultWorkbookCoreReportSupport::toOoxmlSignatureReport)
            .toList());
  }

  static OoxmlEncryptionReport toOoxmlEncryptionReport(
      dev.erst.gridgrind.excel.ExcelOoxmlEncryptionSnapshot snapshot) {
    return new OoxmlEncryptionReport(
        snapshot.encrypted(),
        snapshot.mode(),
        snapshot.cipherAlgorithm(),
        snapshot.hashAlgorithm(),
        snapshot.chainingMode(),
        snapshot.keyBits(),
        snapshot.blockSize(),
        snapshot.spinCount());
  }

  static OoxmlSignatureReport toOoxmlSignatureReport(
      dev.erst.gridgrind.excel.ExcelOoxmlSignatureSnapshot snapshot) {
    return new OoxmlSignatureReport(
        snapshot.packagePartName(),
        snapshot.signerSubject(),
        snapshot.signerIssuer(),
        snapshot.serialNumberHex(),
        snapshot.state());
  }

  static GridGrindResponse.NamedRangeReport toNamedRangeReport(ExcelNamedRangeSnapshot namedRange) {
    return switch (namedRange) {
      case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot ->
          new GridGrindResponse.NamedRangeReport.RangeReport(
              rangeSnapshot.name(),
              toNamedRangeScope(rangeSnapshot.scope()),
              rangeSnapshot.refersToFormula(),
              new NamedRangeTarget(
                  rangeSnapshot.target().sheetName(), rangeSnapshot.target().range()));
      case ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot ->
          new GridGrindResponse.NamedRangeReport.FormulaReport(
              formulaSnapshot.name(),
              toNamedRangeScope(formulaSnapshot.scope()),
              formulaSnapshot.refersToFormula());
    };
  }

  static NamedRangeScope toNamedRangeScope(ExcelNamedRangeScope scope) {
    return switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> new NamedRangeScope.Workbook();
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          new NamedRangeScope.Sheet(sheetScope.sheetName());
    };
  }

  static WorkbookProtectionReport toWorkbookProtectionReport(
      ExcelWorkbookProtectionSnapshot protection) {
    return new WorkbookProtectionReport(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        protection.workbookPasswordHashPresent(),
        protection.revisionsPasswordHashPresent());
  }
}
