package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;

/** Converts workbook-global security, summary, protection, and named-range snapshots. */
final class InspectionResultWorkbookCoreReportSupport {
  private InspectionResultWorkbookCoreReportSupport() {}

  static WorkbookSummary toWorkbookSummary(
      dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary workbookSummary) {
    return switch (workbookSummary) {
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.Empty empty ->
          new WorkbookSummary.Empty(
              empty.sheetCount(),
              empty.sheetNames().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                  .toList(),
              empty.namedRangeCount(),
              empty.forceFormulaRecalculationOnOpen());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.WithSheets withSheets ->
          new WorkbookSummary.WithSheets(
              withSheets.sheetCount(),
              withSheets.sheetNames().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                  .toList(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(withSheets.activeSheetName()),
              withSheets.selectedSheetNames().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                  .toList(),
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
    return switch (snapshot) {
      case dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot.None _ ->
          new OoxmlEncryptionReport.None();
      case dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot.Encrypted encrypted ->
          new OoxmlEncryptionReport.Encrypted(
              encrypted.mode(),
              encrypted.cipherAlgorithm(),
              encrypted.hashAlgorithm(),
              encrypted.chainingMode(),
              encrypted.keyBits(),
              encrypted.blockSize(),
              encrypted.spinCount());
    };
  }

  static OoxmlSignatureReport toOoxmlSignatureReport(
      dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureSnapshot snapshot) {
    return new OoxmlSignatureReport(
        snapshot.packagePartName(),
        snapshot
            .signer()
            .map(
                signer ->
                    new OoxmlSignatureReport.SignerIdentity(
                        signer.subject(), signer.issuer(), signer.serialNumberHex())),
        snapshot.state());
  }
}
