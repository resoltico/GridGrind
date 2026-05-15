package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlDsigProviderSupport;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;

/** Public signature-info bridge for callers that need GridGrind's xmlsec-4-compatible setup. */
public final class WorkbookSignatureSupport {
  private WorkbookSignatureSupport() {}

  /** Returns a POI signature wrapper bound to GridGrind's xmlsec-4-compatible dsig provider. */
  public static SignatureInfo newSignatureInfo() {
    return ExcelOoxmlDsigProviderSupport.newSignatureInfo();
  }
}
