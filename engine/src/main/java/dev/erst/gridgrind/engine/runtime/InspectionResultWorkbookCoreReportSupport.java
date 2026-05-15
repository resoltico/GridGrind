package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.NamedRangeReport;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.dto.WorkbookSummary;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSnapshot;

/** Converts workbook-global security, summary, protection, and named-range snapshots. */
final class InspectionResultWorkbookCoreReportSupport {
  private InspectionResultWorkbookCoreReportSupport() {}

  static WorkbookSummary toWorkbookSummary(
      dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary workbookSummary) {
    return switch (workbookSummary) {
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.Empty empty ->
          new WorkbookSummary.Empty(
              empty.sheetCount(),
              empty.sheetNames(),
              empty.namedRangeCount(),
              empty.forceFormulaRecalculationOnOpen());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.WithSheets withSheets ->
          new WorkbookSummary.WithSheets(
              withSheets.sheetCount(),
              withSheets.sheetNames(),
              withSheets.activeSheetName(),
              withSheets.selectedSheetNames(),
              withSheets.namedRangeCount(),
              withSheets.forceFormulaRecalculationOnOpen());
    };
  }

  static OoxmlPackageSecurityReport toOoxmlPackageSecurityReport(
      dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot snapshot) {
    return new OoxmlPackageSecurityReport(
        toOoxmlEncryptionReport(snapshot.encryption()),
        snapshot.signatures().stream()
            .map(InspectionResultWorkbookCoreReportSupport::toOoxmlSignatureReport)
            .toList());
  }

  static OoxmlEncryptionReport toOoxmlEncryptionReport(
      dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot snapshot) {
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
      dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureSnapshot snapshot) {
    return new OoxmlSignatureReport(
        snapshot.packagePartName(),
        snapshot.signerSubject(),
        snapshot.signerIssuer(),
        snapshot.serialNumberHex(),
        snapshot.state());
  }

  static NamedRangeReport toNamedRangeReport(ExcelNamedRangeSnapshot namedRange) {
    return switch (namedRange) {
      case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot ->
          new NamedRangeReport.RangeReport(
              rangeSnapshot.name(),
              toNamedRangeScope(rangeSnapshot.scope()),
              rangeSnapshot.refersToFormula(),
              NamedRangeTarget.range(
                  ((ExcelNamedRangeTarget.Range) rangeSnapshot.target()).sheetName(),
                  ((ExcelNamedRangeTarget.Range) rangeSnapshot.target()).range()));
      case ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot ->
          new NamedRangeReport.FormulaReport(
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
