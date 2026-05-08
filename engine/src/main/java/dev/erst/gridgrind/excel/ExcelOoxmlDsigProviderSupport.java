package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.spi.ExcelOoxmlXmlSec4RelationshipTransformService;
import java.security.Provider;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;

/** Owns the xmlsec-4-compatible OOXML dsig transform registration used by POI signing flows. */
final class ExcelOoxmlDsigProviderSupport {
  private static final Provider XML_DSIG_PROVIDER = newXmlDsigProvider();

  private ExcelOoxmlDsigProviderSupport() {}

  /** Returns a POI signature wrapper bound to GridGrind's xmlsec-4-compatible dsig provider. */
  static SignatureInfo newSignatureInfo() {
    SignatureInfo signatureInfo = new SignatureInfo();
    signatureInfo.setProvider(XML_DSIG_PROVIDER);
    return signatureInfo;
  }

  static Provider provider() {
    return XML_DSIG_PROVIDER;
  }

  private static Provider newXmlDsigProvider() {
    Provider provider = new org.apache.jcp.xml.dsig.internal.dom.XMLDSigRI();
    provider.put(
        "TransformService." + ExcelOoxmlXmlSec4RelationshipTransformService.TRANSFORM_URI,
        ExcelOoxmlXmlSec4RelationshipTransformService.class.getName());
    provider.put(
        "TransformService."
            + ExcelOoxmlXmlSec4RelationshipTransformService.TRANSFORM_URI
            + " MechanismType",
        "DOM");
    return provider;
  }
}
