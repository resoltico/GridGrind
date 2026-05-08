package dev.erst.gridgrind.excel.spi;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

import java.security.InvalidAlgorithmParameterException;
import java.security.spec.AlgorithmParameterSpec;
import java.util.ArrayList;
import java.util.List;
import javax.xml.crypto.Data;
import javax.xml.crypto.MarshalException;
import javax.xml.crypto.OctetStreamData;
import javax.xml.crypto.XMLCryptoContext;
import javax.xml.crypto.XMLStructure;
import javax.xml.crypto.dom.DOMStructure;
import javax.xml.crypto.dsig.TransformException;
import javax.xml.crypto.dsig.TransformService;
import javax.xml.crypto.dsig.spec.TransformParameterSpec;
import org.apache.jcp.xml.dsig.internal.dom.ApacheNodeSetData;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.apache.xml.security.signature.XMLSignatureNodeInput;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.xpackage.x2006.digitalSignature.CTRelationshipReference;
import org.openxmlformats.schemas.xpackage.x2006.digitalSignature.RelationshipReferenceDocument;
import org.w3.x2000.x09.xmldsig.TransformDocument;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Xmlsec 4 made {@code XMLSignatureInput} abstract, so GridGrind owns the OOXML relationship
 * transform registration POI still assumes is concrete.
 */
public final class ExcelOoxmlXmlSec4RelationshipTransformService extends TransformService {
  public static final String TRANSFORM_URI =
      "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";

  private final List<String> sourceIds = new ArrayList<>();
  private final org.apache.poi.poifs.crypt.dsig.services.RelationshipTransformService
      marshallingDelegate =
          new org.apache.poi.poifs.crypt.dsig.services.RelationshipTransformService();

  /** Creates the provider-instantiated OOXML relationship transform service used by XML DSig. */
  public ExcelOoxmlXmlSec4RelationshipTransformService() {}

  @Override
  public void init(TransformParameterSpec params) throws InvalidAlgorithmParameterException {
    marshallingDelegate.init(params);
  }

  @Override
  public void init(XMLStructure parent, XMLCryptoContext context)
      throws InvalidAlgorithmParameterException {
    DOMStructure domParent = (DOMStructure) parent;
    Node parentNode = domParent.getNode();

    try {
      TransformDocument transformDocument =
          TransformDocument.Factory.parse(parentNode, DEFAULT_XML_OPTIONS);
      XmlObject[] relationshipReferences =
          transformDocument
              .getTransform()
              .selectChildren(RelationshipReferenceDocument.type.getDocumentElementName());
      sourceIds.clear();
      for (XmlObject xmlObject : relationshipReferences) {
        sourceIds.add(((CTRelationshipReference) xmlObject).getSourceId());
      }
    } catch (XmlException exception) {
      throw new InvalidAlgorithmParameterException(exception);
    }
  }

  @Override
  public void marshalParams(XMLStructure parent, XMLCryptoContext context) throws MarshalException {
    marshallingDelegate.marshalParams(parent, context);
    try {
      init(parent, context);
    } catch (InvalidAlgorithmParameterException exception) {
      throw new MarshalException(exception);
    }
  }

  @Override
  public @Nullable AlgorithmParameterSpec getParameterSpec() {
    return null;
  }

  @Override
  public Data transform(Data data, XMLCryptoContext context) throws TransformException {
    OctetStreamData octetStreamData = (OctetStreamData) data;
    Document document;
    try (java.io.InputStream octetStream = octetStreamData.getOctetStream()) {
      document = DocumentHelper.readDocument(octetStream);
    } catch (Exception exception) {
      throw new TransformException(exception.getMessage(), exception);
    }

    Element root = document.getDocumentElement();
    NodeList children = root.getChildNodes();
    List<Element> relationships = new ArrayList<>();
    for (int index = children.getLength() - 1; index >= 0; index--) {
      Node child = children.item(index);
      if (!"Relationship".equals(child.getLocalName())) {
        root.removeChild(child);
        continue;
      }
      Element relationship = (Element) child;
      String relationshipId = relationship.getAttribute("Id");
      if (!sourceIds.contains(relationshipId)) {
        root.removeChild(child);
        continue;
      }
      String targetMode = relationship.getAttribute("TargetMode");
      if (targetMode.isEmpty()) {
        relationship.setAttribute("TargetMode", "Internal");
      }
      relationships.add(relationship);
      root.removeChild(child);
    }

    relationships.sort(
        java.util.Comparator.comparing(relationship -> relationship.getAttribute("Id")));
    for (Element relationship : relationships) {
      root.appendChild(relationship);
    }

    return new ApacheNodeSetData(new XMLSignatureNodeInput(root));
  }

  @Override
  public Data transform(Data data, XMLCryptoContext context, java.io.OutputStream outputStream)
      throws TransformException {
    return transform(data, context);
  }

  @Override
  public boolean isFeatureSupported(String feature) {
    return false;
  }
}
