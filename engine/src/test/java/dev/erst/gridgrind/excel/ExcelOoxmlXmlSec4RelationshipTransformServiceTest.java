package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.spi.ExcelOoxmlXmlSec4RelationshipTransformService;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.InvalidAlgorithmParameterException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import javax.xml.crypto.MarshalException;
import javax.xml.crypto.NodeSetData;
import javax.xml.crypto.OctetStreamData;
import javax.xml.crypto.dom.DOMStructure;
import javax.xml.crypto.dsig.TransformException;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.apache.poi.poifs.crypt.dsig.services.RelationshipTransformService.RelationshipTransformParameterSpec;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

/** Regression coverage for GridGrind's xmlsec-4-compatible OOXML relationship transform. */
class ExcelOoxmlXmlSec4RelationshipTransformServiceTest {
  @Test
  void marshalParamsAndTransformPreserveOnlySelectedRelationshipsInSortedOrder() throws Exception {
    ExcelOoxmlXmlSec4RelationshipTransformService configuredTransformService =
        new ExcelOoxmlXmlSec4RelationshipTransformService();
    RelationshipTransformParameterSpec params = new RelationshipTransformParameterSpec();
    params.addRelationshipReference("rId2");
    params.addRelationshipReference("rId1");
    configuredTransformService.init(params);

    assertNull(configuredTransformService.getParameterSpec());

    Document transformDocument = transformDocument();
    Element transformElement = transformDocument.getDocumentElement();
    configuredTransformService.marshalParams(new DOMStructure(transformElement), null);
    assertEquals(
        2,
        transformElement.getElementsByTagNameNS(OO_DIGSIG_NS, "RelationshipReference").getLength());

    ExcelOoxmlXmlSec4RelationshipTransformService unmarshalledTransformService =
        new ExcelOoxmlXmlSec4RelationshipTransformService();
    unmarshalledTransformService.init(new DOMStructure(transformElement), null);

    NodeSetData transformedRelationships =
        assertInstanceOf(
            NodeSetData.class,
            unmarshalledTransformService.transform(
                relationshipsData(
                    """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId2" Type="type" Target="b"/>
                      <IgnoreMe/>
                      <Relationship Id="skip" Type="type" Target="skip" TargetMode="External"/>
                      <Relationship Id="rId1" Type="type" Target="a" TargetMode="External"/>
                    </Relationships>
                    """),
                null));
    List<Element> relationships = relationshipElements(transformedRelationships);
    assertEquals(List.of("rId1", "rId2"), relationshipIds(relationships));
    assertEquals("External", relationships.get(0).getAttribute("TargetMode"));
    assertEquals("Internal", relationships.get(1).getAttribute("TargetMode"));

    NodeSetData transformedWithOutputStream =
        assertInstanceOf(
            NodeSetData.class,
            unmarshalledTransformService.transform(
                relationshipsData(
                    """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="type" Target="a"/>
                    </Relationships>
                    """),
                null,
                new ByteArrayOutputStream()));
    assertEquals(
        List.of("rId1"), relationshipIds(relationshipElements(transformedWithOutputStream)));
  }

  @Test
  void initRejectsNonTransformStructure() {
    ExcelOoxmlXmlSec4RelationshipTransformService transformService =
        new ExcelOoxmlXmlSec4RelationshipTransformService();

    assertThrows(
        InvalidAlgorithmParameterException.class,
        () ->
            transformService.init(new DOMStructure(invalidDocument().getDocumentElement()), null));
  }

  @Test
  void marshalParamsWrapsInvalidTransformStructures() throws Exception {
    ExcelOoxmlXmlSec4RelationshipTransformService transformService =
        new ExcelOoxmlXmlSec4RelationshipTransformService();
    RelationshipTransformParameterSpec params = new RelationshipTransformParameterSpec();
    params.addRelationshipReference("rId1");
    transformService.init(params);

    assertThrows(
        MarshalException.class,
        () ->
            transformService.marshalParams(
                new DOMStructure(invalidDocument().getDocumentElement()), null));
  }

  @Test
  void transformRejectsInvalidXmlAndReportsUnsupportedFeatures() throws Exception {
    ExcelOoxmlXmlSec4RelationshipTransformService transformService =
        new ExcelOoxmlXmlSec4RelationshipTransformService();
    RelationshipTransformParameterSpec params = new RelationshipTransformParameterSpec();
    params.addRelationshipReference("rId1");
    transformService.init(params);

    assertFalse(transformService.isFeatureSupported("gridgrind.test"));
    assertThrows(
        TransformException.class,
        () ->
            transformService.transform(
                new OctetStreamData(
                    new ByteArrayInputStream("not xml".getBytes(StandardCharsets.UTF_8))),
                null));
  }

  private static Document transformDocument() throws IOException, SAXException {
    return DocumentHelper.readDocument(
        new ByteArrayInputStream(
            """
            <Transform xmlns="http://www.w3.org/2000/09/xmldsig#"/>
            """
                .getBytes(StandardCharsets.UTF_8)));
  }

  private static Document invalidDocument() throws IOException, SAXException {
    return DocumentHelper.readDocument(
        new ByteArrayInputStream("<NotTransform/>".getBytes(StandardCharsets.UTF_8)));
  }

  private static OctetStreamData relationshipsData(String xml) {
    return new OctetStreamData(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));
  }

  private static List<Element> relationshipElements(NodeSetData nodeSetData) {
    List<Element> relationships = new ArrayList<>();
    Iterator<?> nodes = nodeSetData.iterator();
    while (nodes.hasNext()) {
      Object node = nodes.next();
      if (node instanceof Element element && "Relationship".equals(element.getLocalName())) {
        relationships.add(element);
      }
    }
    return relationships;
  }

  private static List<String> relationshipIds(List<Element> relationships) {
    List<String> relationshipIds = new ArrayList<>();
    for (Element relationship : relationships) {
      relationshipIds.add(relationship.getAttribute("Id"));
    }
    return relationshipIds;
  }

  private static final String OO_DIGSIG_NS =
      "http://schemas.openxmlformats.org/package/2006/digital-signature";
}
